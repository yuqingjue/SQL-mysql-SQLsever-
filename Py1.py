import pandas as pd
import json
import dash
from dash import html, dcc, Input, Output, State
import dash_cytoscape as cyto

import pandas as pd
import json
import dash
from dash import html, dcc, Input, Output, State
import dash_cytoscape as cyto


def make_unique_headers(headers):
    """为重复的列名添加数字后缀，返回唯一列名列表"""
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
    """
    从 df 的 start_row 开始读取数据行，直到遇到空行或 stop_marker（如果指定）为止。
    返回 (data_rows, headers) 其中 headers 是 start_row 行的内容，data_rows 是后续行的列表。
    """
    headers = df.iloc[start_row, :].tolist()
    data = []
    for j in range(start_row + 1, len(df)):
        row = df.iloc[j, :]
        if pd.isna(row.iloc[0]) or (stop_marker is not None and row.iloc[0] == stop_marker):
            break
        data.append(row.tolist())
    return data, headers

def aggregate_steps(steps_list):
    """将步骤列表聚合成一个结构化的 JSON 对象字符串（优化版）"""
    if not steps_list:
        return json.dumps({})
    
    def _update_main_source(cur, step):
        """如果当前 main_source 为空，且步骤的 MainSource 不是 '(previous)'，则使用它"""
        if not cur:
            src = step.get('MainSource', '')
            if src and src != '(previous)':
                return src
        return cur

    main_source = None
    pre_filters = []
    joins = []
    window = None
    post_filters = []
    output_fields = None

    for s in steps_list:
        op = s.get('Operation', '').upper()
        if op == 'FILTER':
            if joins or window:
                post_filters.append(s.get('FilterCondition', ''))
            else:
                pre_filters.append(s.get('FilterCondition', ''))
            main_source = _update_main_source(main_source, s)
        elif op == 'JOIN':
            main_source = _update_main_source(main_source, s)
            joins.append({
                "table": s.get('JoinTable', ''),
                "type": s.get('JoinType', ''),
                "on": s.get('JoinCondition', ''),
                "extra_filter": s.get('FilterCondition', '')
            })
        elif op == 'WINDOW':
            window = s.get('WindowFunction', '')
            main_source = _update_main_source(main_source, s)
        elif op in ('SELECT', 'AGGREGATE'):
            output_fields = s.get('OutputFields', '')
            main_source = _update_main_source(main_source, s)

    pre_filter = " AND ".join(pre_filters) if pre_filters else ""
    post_filter = " AND ".join(post_filters) if post_filters else ""

    aggregated = {
        "main_source": main_source,
        "joins": joins,
        "pre_filter": pre_filter,
        "window": window or "",
        "post_filter": post_filter,
        "output_fields": output_fields or ""
    }
    return json.dumps(aggregated, indent=2, ensure_ascii=False)

def make_step(step_num, operation, main_source, join_table="", join_type="",
              join_cond="", filter_cond="", window_func="", output_fields="", etl_partition=""):
    return {
        "Step": step_num,
        "Operation": operation,
        "MainSource": main_source,
        "JoinTable": join_table,
        "JoinType": join_type,
        "JoinCondition": join_cond,
        "FilterCondition": filter_cond,
        "WindowFunction": window_func,
        "ETL_Partition": etl_partition,   # 改为参数
        "OutputFields": output_fields
    }


def disaggregate_steps(aggregated_json, table_name):
    """将聚合 JSON 解析回步骤列表（优化版）"""
    agg = json.loads(aggregated_json)
    steps = []
    step = 1
    main_source = agg.get("main_source", "")
    current_source = main_source if main_source else ""

    # 前置过滤
    if agg.get("pre_filter"):
        steps.append(make_step(step, "FILTER", current_source, filter_cond=agg["pre_filter"]))
        step += 1
        current_source = "(previous)"

    # JOIN
    for j in agg.get("joins", []):
        steps.append(make_step(step, "JOIN", current_source,
                               join_table=j.get("table", ""),
                               join_type=j.get("type", ""),
                               join_cond=j.get("on", ""),
                               filter_cond=j.get("extra_filter", "")))
        step += 1
        current_source = "(previous)"

    # WINDOW
    if agg.get("window"):
        steps.append(make_step(step, "WINDOW", "(previous)", window_func=agg["window"]))
        step += 1
        current_source = "(previous)"

    # 后置过滤
    if agg.get("post_filter"):
        steps.append(make_step(step, "FILTER", "(previous)", filter_cond=agg["post_filter"]))
        step += 1
        current_source = "(previous)"

    # 输出字段
    if agg.get("output_fields"):
        steps.append(make_step(step, "SELECT", "(previous)", output_fields=agg["output_fields"]))
        step += 1

    return steps

# ==================== 1. Read Excel (single table per sheet) ====================
EXCEL_PATH = "/Users/taixujianyi/Documents/Test_file/data_flows.xlsx"

def load_metadata(path):
    xls = pd.ExcelFile(path)
    sheet_names = xls.sheet_names

    tables_data = []
    steps_list = []
    flows_list = []

    for sheet in sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        # 表信息
        info = dict(zip(df.iloc[0, :5].tolist(), df.iloc[1, :5].tolist()))
        tables_data.append(info)

        # 查找步骤表头
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

        # 查找字段映射表头
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
                flows_list.append(flow_df)

    # 合并所有 sheet 的数据
    tables_df = pd.DataFrame(tables_data)
    steps_df = pd.concat(steps_list, ignore_index=True) if steps_list else pd.DataFrame(columns=["TargetTable","Step","Operation","MainSource","JoinTable","JoinType","JoinCondition","FilterCondition","WindowFunction","ETL_Partition","OutputFields"])
    flows_df = pd.concat(flows_list, ignore_index=True) if flows_list else pd.DataFrame(columns=["SourceTable","SourceField","TargetTable","TargetField","Rule","Notes"])

    return tables_df, flows_df, steps_df

tables_df, flows_df, steps_df = load_metadata(EXCEL_PATH)

# ==================== 2. Table-level graph elements ====================
def build_table_elements(tables, flows):
    nodes = [
        {'data': {'id': row['TableName'],
                  'label': row['TableName'],
                  'category': row.get('Category', 'process')}}
        for _, row in tables.iterrows()
    ]
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

# ==================== 3. Column-level graph elements ====================
def build_field_view_elements(tables, flows):
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
        fields = sorted(list(table_fields.get(tbl_name, [])))
        n = len(fields)
        parent_height = header_height + n * field_height + (n - 1) * field_gap + 30
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

        for i, field in enumerate(fields):
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

# ==================== 4. Stylesheets (Macaron blue theme) ====================
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
    {
        'selector': 'edge',
        'style': {
            'target-arrow-shape': 'triangle',
            'curve-style': 'bezier',
            'line-color': '#7f8c8d',
            'target-arrow-color': '#7f8c8d',
        }
    },
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
    {
        'selector': 'edge',
        'style': {
            'width': 1.5,
            'line-color': '#0074D9',
            'target-arrow-shape': 'triangle',
            'target-arrow-color': '#0074D9',
            'font-size': '8px',
            'curve-style': 'bezier',
        }
    },
    {
        'selector': 'edge:selected',
        'style': {
            'line-color': '#0074D9',
            'target-arrow-color': '#0074D9',
            'width': 2,
            'overlay-color': '#00BFFF',
            'overlay-opacity': 0.5,
            'overlay-padding': 4,
        }
    },
]

# ==================== 5. Unused but kept ====================
def render_step(step):
    def safe_str(val):
        return str(val) if pd.notna(val) else ''
    op = safe_str(step.get('Operation')).upper()
    items = [html.P(f"Step {step['Step']}: {op}", style={'fontWeight': 'bold'})]
    if op == 'JOIN':
        items.append(html.P(f"  {safe_str(step.get('MainSource'))} {safe_str(step.get('JoinType'))} JOIN {safe_str(step.get('JoinTable'))}"))
        if step.get('JoinCondition'):
            items.append(html.P(f"  ON {safe_str(step.get('JoinCondition'))}"))
    elif op == 'FILTER':
        items.append(html.P(f"  WHERE {safe_str(step.get('FilterCondition'))}"))
    elif op == 'WINDOW':
        items.append(html.P(f"  {safe_str(step.get('WindowFunction'))}"))
    if step.get('OutputFields'):
        items.append(html.P(f"  Output: {safe_str(step.get('OutputFields'))}"))
    return html.Div(items, style={'marginLeft': 20, 'marginBottom': 15})

# ==================== 6. Save to Excel (new per-sheet format) ====================
def write_to_excel(tables, flows, steps, path):
    """Write back to single-sheet-per-table Excel"""
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet
    
    # Prepare steps per table
    step_groups = {}
    for _, row in steps.iterrows():
        tbl = row['TargetTable']
        step_groups.setdefault(tbl, []).append(row)
    
    # Prepare flows per table
    flow_groups = {}
    for _, row in flows.iterrows():
        # associate with both source and target tables
        for tbl in (row['SourceTable'], row['TargetTable']):
            flow_groups.setdefault(tbl, []).append(row)
        # Actually we need to avoid duplicates, so we'll just collect per table by checking if source or target equals table
        # Better approach: For each table, select relevant flows when writing sheet.
    
    # We will use flows_df directly inside the loop to filter.
    for _, tbl_row in tables.iterrows():
        tbl_name = tbl_row['TableName']
        ws = wb.create_sheet(title=tbl_name)
        
        # 表信息
        info_headers = ["Database", "TableName", "Description", "Category", "UpdateFrequency"]
        ws.append(info_headers)
        ws.append([tbl_row.get('Database',''), tbl_row['TableName'], tbl_row.get('Description',''), tbl_row.get('Category',''), tbl_row.get('UpdateFrequency','')])
        ws.append([])  # empty row
        
        # 生成步骤
        step_headers = ["Step", "Operation", "MainSource", "JoinTable", "JoinType", "JoinCondition",
                        "FilterCondition", "WindowFunction", "ETL_Partition", "OutputFields"]
        ws.append(step_headers)
        tbl_steps = step_groups.get(tbl_name, [])
        for s in tbl_steps:
            ws.append([s.get(h, '') for h in step_headers])
        if not tbl_steps:
            ws.append([''] * len(step_headers))
        ws.append([])  # empty row
        
        # 字段映射
        flow_headers = ["SourceTable", "SourceField", "TargetTable", "TargetField", "Rule", "Notes"]
        ws.append(flow_headers)
        # filter flows related to this table
        tbl_flows = flows[(flows['SourceTable'] == tbl_name) | (flows['TargetTable'] == tbl_name)]
        for _, f in tbl_flows.iterrows():
            ws.append([f.get('SourceTable',''), f.get('SourceField',''), f.get('TargetTable',''), f.get('TargetField',''), f.get('Rule',''), f.get('Notes','')])
    
    wb.save(path)


# ==================== 7. Dash App ====================
app = dash.Dash(__name__)
app.config.suppress_callback_exceptions = True

app.layout = html.Div([
    html.H3("Flip Data Logic - in progress"),

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
        html.H4("Field Generation Logic (click on the connection to view)"),
        html.Div(id='field-info-panel', children="Switch to column-level view and click a connection"),
        html.Div(id='rule-save-status'),
    ]),

    dcc.Store(id='view-mode', data='table'),
    dcc.Store(id='expanded-table', data=None),
    dcc.Store(id='selected-edge', data=None),
    dcc.Store(id='global-dfs', data={
        'tables': tables_df.to_dict('records'),
        'flows': flows_df.to_dict('records'),
        'steps': steps_df.to_dict('records')
    }),
])

# ==================== 8. Callback: Toggle view ====================
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
    State('global-dfs', 'data')
)
def toggle_view(n_clicks, current_mode, dfs_dict):
    if n_clicks == 0:
        raise dash.exceptions.PreventUpdate

    tables = pd.DataFrame(dfs_dict['tables'])
    flows = pd.DataFrame(dfs_dict['flows'])

    if current_mode == 'table':
        elements = build_field_view_elements(tables, flows)
        edit_style = {'display': 'none'}
        field_style = {'display': 'block', 'padding': 10, 'border': '1px solid #ccc', 'borderRadius': 5}
        return (elements, field_stylesheet,
                {'name': 'preset', 'fit': True, 'padding': 60},
                'field', 'Current: Column-level View', 'Switch to Table-level View',
                "Click a connection to view its generation rule",
                '', edit_style, field_style)
    else:
        elements = build_table_elements(tables, flows)
        edit_style = {'display': 'block', 'padding': 10, 'border': '1px solid #ccc', 'borderRadius': 5}
        field_style = {'display': 'none'}
        return (elements, base_stylesheet,
                {'name': 'breadthfirst', 'spacingFactor': 1.2},
                'table', 'Current: Table-level View', 'Switch to Column View',
                "Switch to column-level view to see field details",
                '', edit_style, field_style)

# ==================== 9. Callback: Table-level edit panel ====================
@app.callback(
    Output('edit-panel', 'children'),
    Output('expanded-table', 'data'),
    Input('graph', 'tapNodeData'),
    State('view-mode', 'data'),
    State('global-dfs', 'data')
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
    field_mappings = flows[(flows['SourceTable'] == table) | (flows['TargetTable'] == table)]
    field_mappings_json = field_mappings[['SourceTable','SourceField','TargetTable','TargetField','Rule']].to_dict('records')

    tables = pd.DataFrame(dfs_dict['tables'])
    table_info = tables[tables['TableName'] == table].iloc[0] if not tables[tables['TableName'] == table].empty else None

    return html.Div([
        html.H5(f"Editing table: {table}"),
        html.Div([
            html.Div([
                html.Label("Description"),
                dcc.Input(id='edit-desc', value=table_info.get('Description','') if table_info is not None else '',
                          style={'width': '100%'}),
                html.Br(),
                html.Label("Category"),
                dcc.Input(id='edit-category', value=table_info.get('Category','process') if table_info is not None else 'process',
                          style={'width': '100%'})
            ], style={'flex': '1', 'padding': '10px'}),
            html.Div([
                html.Label("Generation Logic (aggregated JSON, modifiable)"),
                dcc.Textarea(
                    id='edit-steps',
                    value=aggregate_steps(steps_list),   # 聚合显示
                    style={'width': '100%', 'height': 150},
                ),
            ], style={'flex': '1', 'padding': '10px'}),
            html.Div([
                html.Label("Field Mapping (JSON array, modifiable)"),
                dcc.Textarea(
                    id='edit-flows',
                    value=json.dumps(field_mappings_json, indent=2, ensure_ascii=False),
                    style={'width': '100%', 'height': 150},
                ),
            ], style={'flex': '1', 'padding': '10px'}),
        ], style={'display': 'flex', 'flex-wrap': 'wrap'}),
        html.Button("Save Changes to Excel", id='save-btn', n_clicks=0,
                    style={'marginTop': '15px', 'display': 'block'}),
    ]), table

# ==================== 10. Column-level edge selection ====================
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

    edge_id = edge_data.get('id', '')
    source = edge_data.get('source', '')
    target = edge_data.get('target', '')
    rule = edge_data.get('label', 'Directly Pull')

    selected = {
        'id': edge_id,
        'source': source,
        'target': target,
        'rule': rule
    }

    children = html.Div([
        html.P(f"From {source} to {target}"),
        html.Label("Generation Rule:"),
        dcc.Input(id='edit-rule-input', value=rule, style={'width': '100%'}),
        html.Button("Save Rule to Excel", id='save-rule-btn', n_clicks=0, style={'marginTop': '10px'}),
    ])
    return children, selected

# ==================== 11. Save rule (column-level) ====================
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

        new_dfs = {
            'tables': tables.to_dict('records'),
            'flows': flows.to_dict('records'),
            'steps': steps.to_dict('records')
        }

        elements = build_field_view_elements(tables, flows)
        return 'Rule saved!', new_dfs, elements, field_stylesheet, {'name': 'preset', 'fit': True, 'padding': 60}

    except Exception as e:
        return f"Save failed: {str(e)}", dash.no_update, dash.no_update, dash.no_update, dash.no_update

# ==================== 12. Save table-level edits ====================
@app.callback(
    Output('save-status', 'children'),
    Output('global-dfs', 'data', allow_duplicate=True),
    Output('graph', 'elements', allow_duplicate=True),
    Output('graph', 'stylesheet', allow_duplicate=True),
    Output('graph', 'layout', allow_duplicate=True),
    Input('save-btn', 'n_clicks'),
    State('edit-desc', 'value'),
    State('edit-category', 'value'),
    State('edit-steps', 'value'),
    State('edit-flows', 'value'),
    State('global-dfs', 'data'),
    State('expanded-table', 'data'),
    State('view-mode', 'data'),
    prevent_initial_call=True
)
def save_table_changes(n_clicks, desc, category, steps_json, flows_json, dfs_dict, expanded_table, view_mode):
    if n_clicks == 0:
        raise dash.exceptions.PreventUpdate
    if view_mode != 'table' or not expanded_table:
        raise dash.exceptions.PreventUpdate

    tables = pd.DataFrame(dfs_dict['tables'])
    flows = pd.DataFrame(dfs_dict['flows'])
    steps = pd.DataFrame(dfs_dict['steps'])
    table = expanded_table

    try:
        # 更新表信息
        mask = tables['TableName'] == table
        if mask.any():
            tables.loc[mask, 'Description'] = desc
            tables.loc[mask, 'Category'] = category

        # 解析聚合步骤为列表
        steps_new = disaggregate_steps(steps_json, table)
        steps = steps[steps['TargetTable'] != table]
        if steps_new:
            df_steps_new = pd.DataFrame(steps_new)
            df_steps_new['TargetTable'] = table
            # 确保列顺序与加载时一致
            cols = ["TargetTable","Step","Operation","MainSource","JoinTable","JoinType","JoinCondition","FilterCondition","WindowFunction","ETL_Partition","OutputFields"]
            df_steps_new = df_steps_new[cols]
            steps = pd.concat([steps, df_steps_new], ignore_index=True)

        # 更新字段映射
        flows_new = json.loads(flows_json)
        flows = flows[(flows['SourceTable'] != table) & (flows['TargetTable'] != table)]
        if flows_new:
            df_flows_new = pd.DataFrame(flows_new)
            flows = pd.concat([flows, df_flows_new], ignore_index=True)

        # 写回 Excel
        write_to_excel(tables, flows, steps, EXCEL_PATH)

        new_dfs = {
            'tables': tables.to_dict('records'),
            'flows': flows.to_dict('records'),
            'steps': steps.to_dict('records')
        }
        global tables_df, flows_df, steps_df
        tables_df = tables
        flows_df = flows
        steps_df = steps

        elements = build_table_elements(tables, flows)
        return "Save successful!", new_dfs, elements, base_stylesheet, {'name': 'breadthfirst'}

    except Exception as e:
        return f"Save failed: {str(e)}", dash.no_update, dash.no_update, dash.no_update, dash.no_update


if __name__ == '__main__':
    app.run(debug=True)