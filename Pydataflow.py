import pandas as pd
import json
import dash
from dash import html, dcc, Input, Output, State,dash_table
import dash_cytoscape as cyto


# ==================== 辅助函数 ====================
def make_unique_headers(headers):
    seen = {}
    unique = []
    for col in headers:
        if col in seen:
            seen[col] += 1
            unique.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            unique.append(col)
    return unique

def extract_data_block(df, start_row, stop_marker=None):
    headers = df.iloc[start_row, :].tolist()
    data = []
    for j in range(start_row + 1, len(df)):
        row = df.iloc[j, :]
        if pd.isna(row.iloc[0]) or (stop_marker is not None and row.iloc[0] == stop_marker):
            break
        data.append(row.tolist())
    return data, headers

# ==================== 读取 Excel（简化的单表格式） ====================
EXCEL_PATH = "/Users/taixujianyi/Documents/Test_file/data_flows.xlsx"

def load_metadata(path):
    xls = pd.ExcelFile(path)
    sheet_names = xls.sheet_names
    tables_data = []
    steps_list = []
    flows_list = []

    for sheet in sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        info = dict(zip(df.iloc[0, :5].tolist(), df.iloc[1, :5].tolist()))
        tables_data.append(info)

        # 步骤
        step_start = None
        for i in range(len(df)):
            if df.iloc[i, 0] == "Step":
                step_start = i
                break
        if step_start is not None:
            step_data, step_headers = extract_data_block(df, step_start, stop_marker="SourceTable")
            step_headers = make_unique_headers(step_headers)
            if step_data:
                step_df = pd.DataFrame(step_data, columns=step_headers)
                step_df.insert(0, 'TargetTable', info['TableName'])
                steps_list.append(step_df)

        # 字段映射（无 TargetTable）
        flow_start = None
        for i in range(len(df)):
            if df.iloc[i, 0] == "SourceTable":
                flow_start = i
                break
        if flow_start is not None:
            flow_data, flow_headers = extract_data_block(df, flow_start)
            flow_headers = make_unique_headers(flow_headers)
            if flow_data:
                flow_df = pd.DataFrame(flow_data, columns=flow_headers)
                flow_df.insert(0, 'TargetTable', info['TableName'])
                flows_list.append(flow_df)

    tables_df = pd.DataFrame(tables_data)
    steps_df = pd.concat(steps_list, ignore_index=True) if steps_list else pd.DataFrame(columns=["TargetTable","Step","Operation","MainSource","JoinTable","JoinType","JoinCondition","FilterCondition","WindowFunction","ETL_Partition","OutputFields","Alias"])
    flows_df = pd.concat(flows_list, ignore_index=True) if flows_list else pd.DataFrame(columns=["SourceTable","SourceField","TargetTable","TargetField","Rule","Notes"])

    # 双向补全
    complementary = []
    for _, row in flows_df.iterrows():
        rev = flows_df[(flows_df['SourceTable'] == row['TargetTable']) & 
                       (flows_df['SourceField'] == row['TargetField']) &
                       (flows_df['TargetTable'] == row['SourceTable']) &
                       (flows_df['TargetField'] == row['SourceField'])]
        if rev.empty:
            complementary.append({
                'SourceTable': row['TargetTable'],
                'SourceField': row['TargetField'],
                'TargetTable': row['SourceTable'],
                'TargetField': row['SourceField'],
                'Rule': row.get('Rule', ''),
                'Notes': row.get('Notes', '')
            })
    if complementary:
        flows_df = pd.concat([flows_df, pd.DataFrame(complementary)], ignore_index=True)

    return tables_df, flows_df, steps_df

tables_df, flows_df, steps_df = load_metadata(EXCEL_PATH)

# ==================== 表级图元素 ====================
def build_table_elements(tables, flows):
    nodes = [{'data': {'id': row['TableName'], 'label': row['TableName'], 'category': row.get('Category', 'process')}} for _, row in tables.iterrows()]
    edges = []
    seen = set()
    for _, row in flows.iterrows():
        pair = (row['SourceTable'], row['TargetTable'])
        if pair not in seen:
            seen.add(pair)
            edges.append({'data': {'source': pair[0], 'target': pair[1]}})
    return nodes + edges

def build_step_dict(steps_df):
    step_dict = {}
    for table, grp in steps_df.groupby('TargetTable'):
        step_dict[table] = grp.sort_values('Step').to_dict('records')
    return step_dict

step_dict = build_step_dict(steps_df)

# ==================== 列级图元素（字段折叠） ====================
MAX_VISIBLE_FIELDS = 8

def build_field_view_elements(tables, flows, expansions=None):
    if expansions is None:
        expansions = {}
    table_fields = {}
    for _, row in flows.iterrows():
        src_tbl, tgt_tbl = row['SourceTable'], row['TargetTable']
        src_f, tgt_f = row['SourceField'], row['TargetField']
        table_fields.setdefault(src_tbl, set()).add(src_f)
        table_fields.setdefault(tgt_tbl, set()).add(tgt_f)
    for _, row in tables.iterrows():
        if row['TableName'] not in table_fields:
            table_fields[row['TableName']] = set()

    col_width = 150
    col_gap = 40
    field_height = 28
    field_gap = 8
    start_x = 40
    start_y = 60
    header_height = 80

    parent_nodes = []
    field_nodes = []
    field_edges = []
    x_cursor = start_x

    for tbl_name in tables['TableName']:
        all_fields = sorted(list(table_fields.get(tbl_name, [])))
        total = len(all_fields)
        expanded = expansions.get(tbl_name, False)
        if expanded or total <= MAX_VISIBLE_FIELDS:
            visible_fields = all_fields
            show_more = False
        else:
            visible_fields = all_fields[:MAX_VISIBLE_FIELDS]
            show_more = True

        n = len(visible_fields)
        extra_height = 30 if show_more else 0
        parent_height = header_height + n * field_height + (n - 1) * field_gap + 30 + extra_height
        parent_width = col_width

        parent_nodes.append({
            'data': {
                'id': tbl_name,
                'label': tbl_name,
                'category': tables[tables['TableName'] == tbl_name]['Category'].values[0],
                'width': parent_width,
                'height': parent_height,
            },
            'position': {'x': x_cursor, 'y': start_y}
        })

        for i, field in enumerate(visible_fields):
            rel_x = 10
            rel_y = header_height + i * (field_height + field_gap)
            field_nodes.append({
                'data': {
                    'id': f"{tbl_name}.{field}",
                    'label': field,
                    'parent': tbl_name
                },
                'position': {'x': rel_x, 'y': rel_y}
            })

        if show_more:
            more_id = f"{tbl_name}.__more__"
            more_y = header_height + n * (field_height + field_gap) + 5
            field_nodes.append({
                'data': {
                    'id': more_id,
                    'label': f'… ({total - MAX_VISIBLE_FIELDS} more)',
                    'parent': tbl_name,
                    'type': 'more-button'
                },
                'position': {'x': 10, 'y': more_y},
                'style': {'font-size': '10px', 'color': '#0074D9', 'text-decoration': 'underline'}
            })

        x_cursor += col_width + col_gap

    for _, row in flows.iterrows():
        src_id = f"{row['SourceTable']}.{row['SourceField']}"
        tgt_id = f"{row['TargetTable']}.{row['TargetField']}"
        field_edges.append({
            'data': {
                'id': f"{src_id}_to_{tgt_id}",
                'source': src_id,
                'target': tgt_id,
                'label': row.get('Rule', '')
            }
        })

    return parent_nodes + field_nodes + field_edges

# ==================== 样式表 ====================
base_stylesheet = [
    {
        'selector': 'node',
        'style': {
            'label': 'data(label)',
            'shape': 'round-rectangle',
            'width': 'label',
            'height': 'label',
            'padding': '12px',
            'text-valign': 'center',
            'text-halign': 'center',
            'font-size': '14px',
            'font-weight': 'bold',
            'color': '#2c3e50',
            'text-wrap': 'wrap',
            'text-max-width': '180px',
            'border-width': 0,
            'background-color': '#e8f4f8',
        }
    },
    {'selector': '[category = "source"]', 'style': {'background-color': '#AEE2FF'}},
    {'selector': '[category = "process"]', 'style': {'background-color': '#B5EAD7'}},
    {'selector': '[category = "datastore"]', 'style': {'background-color': '#C3E0E5'}},
    {'selector': '[category = "destination"]', 'style': {'background-color': '#A9C6D9'}},
    {'selector': 'edge', 'style': {'target-arrow-shape': 'triangle', 'curve-style': 'bezier', 'line-color': '#7f8c8d', 'target-arrow-color': '#7f8c8d'}},
]

field_stylesheet = [
    {
        'selector': 'node',
        'style': {
            'label': 'data(label)',
            'shape': 'round-rectangle',
            'width': 'data(width)',
            'height': 'data(height)',
            'padding': '5px',
            'text-valign': 'top',
            'text-halign': 'center',
            'font-size': '16px',
            'font-weight': 'bold',
            'color': '#2c3e50',
            'border-width': 0,
            'background-color': '#e8f4f8',
        }
    },
    {'selector': '[category = "source"]', 'style': {'background-color': '#AEE2FF'}},
    {'selector': '[category = "process"]', 'style': {'background-color': '#B5EAD7'}},
    {'selector': '[category = "datastore"]', 'style': {'background-color': '#C3E0E5'}},
    {'selector': '[category = "destination"]', 'style': {'background-color': '#A9C6D9'}},
    {
        'selector': 'node[parent]',
        'style': {
            'shape': 'round-rectangle',
            'width': 130,
            'height': 28,
            'background-color': '#ffffff',
            'border-width': 1,
            'border-color': '#A9CCE3',
            'font-size': '12px',
            'text-valign': 'center',
            'text-halign': 'center',
            'label': 'data(label)',
            'color': '#2c3e50',
        }
    },
    {'selector': 'edge', 'style': {'width': 1.5, 'line-color': '#0074D9', 'target-arrow-shape': 'triangle', 'target-arrow-color': '#0074D9', 'font-size': '8px', 'curve-style': 'bezier'}},
    {'selector': 'edge:selected', 'style': {'line-color': '#0074D9', 'target-arrow-color': '#0074D9', 'width': 2, 'overlay-color': '#00BFFF', 'overlay-opacity': 0.5, 'overlay-padding': 4}},
]

# ==================== 保存到 Excel ====================
def write_to_excel(tables, flows, steps, path):
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    step_groups = {}
    for _, row in steps.iterrows():
        tbl = row['TargetTable']
        step_groups.setdefault(tbl, []).append(row)
    for _, tbl_row in tables.iterrows():
        tbl_name = tbl_row['TableName']
        ws = wb.create_sheet(title=tbl_name)
        info_headers = ["Database", "TableName", "Description", "Category", "UpdateFrequency"]
        ws.append(info_headers)
        ws.append([tbl_row.get('Database',''), tbl_row['TableName'], tbl_row.get('Description',''), tbl_row.get('Category',''), tbl_row.get('UpdateFrequency','')])
        ws.append([])
        step_headers = ["Step", "Operation", "MainSource", "JoinTable", "JoinType", "JoinCondition",
                        "FilterCondition", "WindowFunction", "ETL_Partition", "OutputFields", "Alias"]
        ws.append(step_headers)
        tbl_steps = step_groups.get(tbl_name, [])
        for s in tbl_steps:
            ws.append([s.get(h, '') for h in step_headers])
        if not tbl_steps:
            ws.append([''] * len(step_headers))
        ws.append([])
        flow_headers = ["SourceTable", "SourceField", "TargetField", "Rule", "Notes"]
        ws.append(flow_headers)
        tbl_flows = flows[flows['TargetTable'] == tbl_name]
        for _, f in tbl_flows.iterrows():
            ws.append([f.get('SourceTable',''), f.get('SourceField',''), f.get('TargetField',''), f.get('Rule',''), f.get('Notes','')])
    wb.save(path)

# ==================== Dash App ====================
app = dash.Dash(__name__)
app.config.suppress_callback_exceptions = True

app.layout = html.Div([
    html.H3("Data Lineage - Editor"),
    html.Div([
        cyto.Cytoscape(
            id='graph',
            layout={'name': 'breadthfirst', 'spacingFactor': 1.2},
            style={'width': '100%', 'height': '600px', 'border': '1px solid #ccc'},
            elements=build_table_elements(tables_df, flows_df),
            stylesheet=base_stylesheet,
            minZoom=0.5,
            maxZoom=2,
        ),
    ]),
    html.Div([
        html.Button("Switch to Column View", id='toggle-view-btn', n_clicks=0),
        html.Span(id='view-mode-indicator', style={'marginLeft': 10}),
    ], style={'marginTop': 10}),
    html.Hr(),
    html.Div(id='edit-section', style={'display': 'block', 'padding': 10, 'border': '1px solid #ccc', 'borderRadius': 5}, children=[
        html.H4("Edit Metadata (Table-level)"),
        html.Div(id='edit-panel', children="Click a table node to edit"),
        html.Div(id='save-status'),
    ]),
    html.Div(id='field-info-section', style={'display': 'none', 'padding': 10, 'border': '1px solid #ccc', 'borderRadius': 5}, children=[
        html.H4("Field Generation Logic (click on the connection/column name to view)"),
        html.Div(id='field-info-panel', children="Switch to column-level view and click a connection"),
        html.Div(id='rule-save-status'),
    ]),
    dcc.Store(id='view-mode', data='table'),
    dcc.Store(id='expanded-table', data=None),
    dcc.Store(id='selected-edge', data=None),
    dcc.Store(id='global-dfs', data={'tables': tables_df.to_dict('records'), 'flows': flows_df.to_dict('records'), 'steps': steps_df.to_dict('records')}),
    dcc.Store(id='field-expansions', data={}),
])

# ==================== 回调：切换视图 ====================
@app.callback(
    Output('graph', 'elements'),
    Output('graph', 'stylesheet'),
    Output('graph', 'layout'),
    Output('view-mode', 'data'),
    Output('view-mode-indicator', 'children'),
    Output('toggle-view-btn', 'children'),
    Output('field-info-panel', 'children'),
    Output('rule-save-status', 'children'),
    Output('edit-section', 'style'),
    Output('field-info-section', 'style'),
    Input('toggle-view-btn', 'n_clicks'),
    State('view-mode', 'data'),
    State('global-dfs', 'data'),
    State('field-expansions', 'data'),
)
def toggle_view(n_clicks, current_mode, dfs_dict, expansions):
    if n_clicks == 0:
        raise dash.exceptions.PreventUpdate
    tables = pd.DataFrame(dfs_dict['tables'])
    flows = pd.DataFrame(dfs_dict['flows'])
    if current_mode == 'table':
        elements = build_field_view_elements(tables, flows, expansions)
        edit_style = {'display': 'none'}
        field_style = {'display': 'block', 'padding': 10, 'border': '1px solid #ccc', 'borderRadius': 5}
        return (elements, field_stylesheet, {'name': 'preset', 'fit': True, 'padding': 60},
                'field', 'Current: Column-level View', 'Switch to Table-level View',
                "Click a connection to view its generation rule", '', edit_style, field_style)
    else:
        elements = build_table_elements(tables, flows)
        edit_style = {'display': 'block', 'padding': 10, 'border': '1px solid #ccc', 'borderRadius': 5}
        field_style = {'display': 'none'}
        return (elements, base_stylesheet, {'name': 'breadthfirst', 'spacingFactor': 1.2},
                'table', 'Current: Table-level View', 'Switch to Column View',
                "Switch to column-level view to see field details", '', edit_style, field_style)

# ==================== 表级编辑面板 ====================
@app.callback(
    Output('edit-panel', 'children'),
    Output('expanded-table', 'data'),
    Input('graph', 'tapNodeData'),
    State('view-mode', 'data'),
    State('global-dfs', 'data'),
)
def show_edit_panel(node_data, view_mode, dfs_dict):
    if view_mode != 'table':
        return html.P("Switch to table-level view to edit"), dash.no_update
    if not node_data:
        return html.P("Click a table node to edit"), None
    table = node_data['id']
    if '.' in table:
        return html.P("Click a table node to edit"), None

    steps = pd.DataFrame(dfs_dict['steps'])
    flows = pd.DataFrame(dfs_dict['flows'])
    step_dict_local = build_step_dict(steps)
    steps_list = step_dict_local.get(table, [])

    step_columns = [
        {'name': 'Step', 'id': 'Step', 'editable': True},
        {'name': 'Operation', 'id': 'Operation', 'editable': True},
        {'name': 'MainSource', 'id': 'MainSource', 'editable': True},
        {'name': 'JoinTable', 'id': 'JoinTable', 'editable': True},
        {'name': 'JoinType', 'id': 'JoinType', 'editable': True},
        {'name': 'JoinCondition', 'id': 'JoinCondition', 'editable': True},
        {'name': 'FilterCondition', 'id': 'FilterCondition', 'editable': True},
        {'name': 'WindowFunction', 'id': 'WindowFunction', 'editable': True},
        {'name': 'OutputFields', 'id': 'OutputFields', 'editable': True},
        {'name': 'Alias', 'id': 'Alias', 'editable': True},
    ]

    step_table = dash_table.DataTable(
        id='edit-steps-table',
        columns=step_columns,
        data=steps_list,
        editable=True,
        row_deletable=True,
        style_table={'overflowX': 'auto', 'width': '100%', 'maxWidth': '100%'},
        style_cell={'textAlign': 'left', 'whiteSpace': 'normal','padding': '5px', 'fontSize': '12px', 'minWidth': '80px'},
        style_header={'backgroundColor': '#f9f9f9', 'fontWeight': 'bold','whiteSpace': 'normal'},
        css=[{'selector': '.dash-spreadsheet td', 'rule': 'border: 1px solid #eee;'}],
    )

    field_columns = [
        {'name': 'Source Table', 'id': 'SourceTable', 'editable': True},
        {'name': 'Source Field', 'id': 'SourceField', 'editable': True},
        {'name': 'Target Field', 'id': 'TargetField', 'editable': True},
        {'name': 'Rule', 'id': 'Rule', 'editable': True},
        {'name': 'Notes', 'id': 'Notes', 'editable': True},
    ]
    field_mappings_df = flows[flows['TargetTable'] == table]
    field_data = field_mappings_df[['SourceTable','SourceField','TargetField','Rule','Notes']].to_dict('records')
    field_table = dash_table.DataTable(
        id='edit-flows-table',
        columns=field_columns,
        data=field_data,
        editable=True,
        row_deletable=True,
        style_table={'overflowX': 'auto', 'maxWidth': '100%'},
        style_cell={'textAlign': 'left', 'padding': '5px', 'fontSize': '12px'},
        style_header={'backgroundColor': '#f9f9f9', 'fontWeight': 'bold'},
        css=[{'selector': '.dash-spreadsheet td', 'rule': 'border: 1px solid #eee;'}],
    )

    tables = pd.DataFrame(dfs_dict['tables'])
    table_info = tables[tables['TableName'] == table].iloc[0] if not tables[tables['TableName'] == table].empty else None

    return html.Div([
        html.H5(f"Editing table: {table}"),
        html.Div([
            html.Div([
                html.Label("Description"),
                dcc.Input(id='edit-desc', value=table_info.get('Description','') if table_info is not None else '', style={'width': '100%'}),
                html.Br(),
                html.Label("Category"),
                dcc.Input(id='edit-category', value=table_info.get('Category','process') if table_info is not None else 'process', style={'width': '100%'})
            ], style={'flex': '1', 'padding': '10px'}),
            html.Div([
                html.Label("Generation Logic (steps table)"),
                step_table
            ], style={'flex': '2', 'padding': '10px'}),
            html.Div([
                html.Label("Field Mapping"),
                field_table
            ], style={'flex': '3', 'padding': '10px'}),
        ], style={'display': 'flex', 'flex-wrap': 'wrap'}),
        html.Button("Save Changes to Excel", id='save-btn', n_clicks=0, style={'marginTop': '15px', 'display': 'block'}),
    ]), table

# ==================== 列级视图点击连线 ====================
@app.callback(
    Output('field-info-panel', 'children', allow_duplicate=True),
    Output('selected-edge', 'data'),
    Input('graph', 'tapEdgeData'),
    State('view-mode', 'data'),
    prevent_initial_call=True
)
def show_field_rule(edge_data, view_mode):
    if view_mode != 'field':
        raise dash.exceptions.PreventUpdate
    if not edge_data:
        return "Click a connection to see its rule", None
    source = edge_data.get('source', '')
    target = edge_data.get('target', '')
    rule = edge_data.get('label', 'Directly Pull')

    tgt_parts = target.split('.')
    tgt_table = tgt_parts[0] if len(tgt_parts) > 1 else ''
    tgt_field = tgt_parts[1] if len(tgt_parts) > 1 else ''

    all_mappings = flows_df[(flows_df['TargetTable'] == tgt_table) & (flows_df['TargetField'] == tgt_field)]
    unique_sources = set()
    for _, row in all_mappings.iterrows():
        unique_sources.add(f"{row['SourceTable']}.{row['SourceField']}")
    source_items = [html.Li(src) for src in sorted(unique_sources)] if unique_sources else [html.Li("No source mapping found")]

    children = html.Div([
        html.P(f"Target Column: {target}"),
        html.P("Source Columns:"),
        html.Ul(source_items),
        html.Hr(),
        html.Label("Edit Rule for this connection:"),
        dcc.Input(id='edit-rule-input', value=rule, style={'width': '100%'}),
        html.Button("Save Rule to Excel", id='save-rule-btn', n_clicks=0, style={'marginTop': '10px'}),
    ])
    selected = {'id': edge_data.get('id', ''), 'source': source, 'target': target, 'rule': rule}
    return children, selected
# ==================== 列级视图点击字段节点显示来源 ====================
@app.callback(
    Output('field-info-panel', 'children', allow_duplicate=True),
    Output('selected-edge', 'data', allow_duplicate=True),  # 清空可能选中的边
    Input('graph', 'tapNodeData'),
    State('view-mode', 'data'),
    prevent_initial_call=True
)
def show_field_source(node_data, view_mode):
    if view_mode != 'field' or not node_data:
        raise dash.exceptions.PreventUpdate
    node_id = node_data.get('id', '')
    # 只处理字段节点（id 包含 '.' 且不是 '...more' 按钮）
    if '.' not in node_id or node_id.endswith('.__more__'):
        raise dash.exceptions.PreventUpdate

    parts = node_id.split('.')
    tgt_table = parts[0]
    tgt_field = parts[1]
    target = f"{tgt_table}.{tgt_field}"

    # 查询该目标字段的所有源映射
    all_mappings = flows_df[(flows_df['TargetTable'] == tgt_table) & (flows_df['TargetField'] == tgt_field)]
    unique_sources = set()
    for _, row in all_mappings.iterrows():
        unique_sources.add(f"{row['SourceTable']}.{row['SourceField']}")
    source_items = [html.Li(src) for src in sorted(unique_sources)] if unique_sources else [html.Li("No source mapping found")]

    children = html.Div([
        html.P(f"Target Column: {target}"),
        html.P("Source Columns:"),
        html.Ul(source_items),
        html.Hr(),
        html.P("Click a connection line to edit a rule."),
    ])
    return children, None   # 清空 selected-edge

# ==================== 保存单字段规则 ====================
@app.callback(
    Output('rule-save-status', 'children', allow_duplicate=True),
    Output('global-dfs', 'data', allow_duplicate=True),
    Output('graph', 'elements', allow_duplicate=True),
    Output('graph', 'stylesheet', allow_duplicate=True),
    Output('graph', 'layout', allow_duplicate=True),
    Input('save-rule-btn', 'n_clicks'),
    State('edit-rule-input', 'value'),
    State('selected-edge', 'data'),
    State('global-dfs', 'data'),
    State('view-mode', 'data'),
    prevent_initial_call=True
)
def save_field_rule(n_clicks, new_rule, selected_edge, dfs_dict, view_mode):
    if n_clicks == 0 or not selected_edge or view_mode != 'field':
        raise dash.exceptions.PreventUpdate
    tables = pd.DataFrame(dfs_dict['tables'])
    flows = pd.DataFrame(dfs_dict['flows'])
    steps = pd.DataFrame(dfs_dict['steps'])
    source = selected_edge['source']
    target = selected_edge['target']
    src_tbl, src_field = source.split('.')
    tgt_tbl, tgt_field = target.split('.')
    try:
        mask = (flows['SourceTable'] == src_tbl) & (flows['SourceField'] == src_field) & \
               (flows['TargetTable'] == tgt_tbl) & (flows['TargetField'] == tgt_field)
        if mask.any():
            flows.loc[mask, 'Rule'] = new_rule
        else:
            return "Error: Mapping not found", dash.no_update, dash.no_update, dash.no_update, dash.no_update
        write_to_excel(tables, flows, steps, EXCEL_PATH)
        global tables_df, flows_df, steps_df
        tables_df = tables
        flows_df = flows
        steps_df = steps
        new_dfs = {'tables': tables.to_dict('records'), 'flows': flows.to_dict('records'), 'steps': steps.to_dict('records')}
        elements = build_field_view_elements(tables, flows)
        return 'Rule saved!', new_dfs, elements, field_stylesheet, {'name': 'preset', 'fit': True, 'padding': 60}
    except Exception as e:
        return f"Save failed: {str(e)}", dash.no_update, dash.no_update, dash.no_update, dash.no_update

# ==================== 保存表级编辑 ====================
@app.callback(
    Output('save-status', 'children'),
    Output('global-dfs', 'data', allow_duplicate=True),
    Output('graph', 'elements', allow_duplicate=True),
    Output('graph', 'stylesheet', allow_duplicate=True),
    Output('graph', 'layout', allow_duplicate=True),
    Input('save-btn', 'n_clicks'),
    State('edit-desc', 'value'),
    State('edit-category', 'value'),
    State('edit-steps-table', 'data'),
    State('edit-flows-table', 'data'),
    State('global-dfs', 'data'),
    State('expanded-table', 'data'),
    State('view-mode', 'data'),
    prevent_initial_call=True
)
def save_table_changes(n_clicks, desc, category, steps_data, flows_data, dfs_dict, expanded_table, view_mode):
    if n_clicks == 0 or view_mode != 'table' or not expanded_table:
        raise dash.exceptions.PreventUpdate
    tables = pd.DataFrame(dfs_dict['tables'])
    flows = pd.DataFrame(dfs_dict['flows'])
    steps = pd.DataFrame(dfs_dict['steps'])
    table = expanded_table
    try:
        mask = tables['TableName'] == table
        if mask.any():
            tables.loc[mask, 'Description'] = desc
            tables.loc[mask, 'Category'] = category

        if steps_data:
            df_steps_new = pd.DataFrame(steps_data)
            for col in ["Step", "Operation", "MainSource", "JoinTable", "JoinType", "JoinCondition",
                        "FilterCondition", "WindowFunction", "OutputFields", "Alias"]:
                if col not in df_steps_new.columns:
                    df_steps_new[col] = ''
            df_steps_new['TargetTable'] = table
        else:
            df_steps_new = pd.DataFrame(columns=["TargetTable","Step","Operation","MainSource","JoinTable","JoinType","JoinCondition","FilterCondition","WindowFunction","OutputFields","Alias"])
        steps = steps[steps['TargetTable'] != table]
        steps = pd.concat([steps, df_steps_new], ignore_index=True)

        if flows_data:
            flows_new_df = pd.DataFrame(flows_data)
            flows_new_df['TargetTable'] = table
            for col in ['SourceTable','SourceField','TargetField','Rule','Notes']:
                if col not in flows_new_df.columns:
                    flows_new_df[col] = ''
            flows_new_df = flows_new_df.dropna(subset=['SourceTable', 'TargetField'])
        else:
            flows_new_df = pd.DataFrame(columns=['SourceTable','SourceField','TargetField','Rule','Notes','TargetTable'])
        flows = flows[flows['TargetTable'] != table]
        flows = pd.concat([flows, flows_new_df], ignore_index=True)

        write_to_excel(tables, flows, steps, EXCEL_PATH)
        new_dfs = {'tables': tables.to_dict('records'), 'flows': flows.to_dict('records'), 'steps': steps.to_dict('records')}
        global tables_df, flows_df, steps_df
        tables_df = tables
        flows_df = flows
        steps_df = steps
        elements = build_table_elements(tables, flows)
        return "Save successful!", new_dfs, elements, base_stylesheet, {'name': 'breadthfirst'}
    except Exception as e:
        return f"Save failed: {str(e)}", dash.no_update, dash.no_update, dash.no_update, dash.no_update

# ==================== 列级视图展开更多字段 ====================
@app.callback(
    Output('field-expansions', 'data'),
    Output('graph', 'elements', allow_duplicate=True),
    Output('graph', 'layout', allow_duplicate=True),
    Input('graph', 'tapNodeData'),
    State('view-mode', 'data'),
    State('field-expansions', 'data'),
    State('global-dfs', 'data'),
    prevent_initial_call=True
)
def expand_table(node_data, view_mode, expansions, dfs_dict):
    if view_mode != 'field' or not node_data:
        raise dash.exceptions.PreventUpdate
    node_id = node_data['id']
    if not node_id.endswith('.__more__'):
        raise dash.exceptions.PreventUpdate
    table_name = node_id.split('.')[0]
    expansions = expansions.copy() if expansions else {}
    expansions[table_name] = True
    tables = pd.DataFrame(dfs_dict['tables'])
    flows = pd.DataFrame(dfs_dict['flows'])
    elements = build_field_view_elements(tables, flows, expansions)
    return expansions, elements, {'name': 'preset', 'fit': True, 'padding': 60}

if __name__ == '__main__':
    app.run(debug=True)