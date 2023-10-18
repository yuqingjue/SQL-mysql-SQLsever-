# -*- coding: utf-8 -*-
"""
Created on Sun Jan  9 18:31:07 2022

@author: Lenovo
"""
import torch
import torch.nn as nn
import tez
import transformers
from transformers import AdamW,get_linear_schedule_with_warmup
from sklearn import metrics
import pandas as pd

class Bertdataset:
    def __init__(self,texts,targets,max_len=64):
        self.targets=targets
        self.texts=texts
        self.tokenizer=transformers.BertTokenizer.from_pretrained(
            "bert-base-uncased",
            do_lower_case=False
            )
        self.max_len=max_len
    def __len__(self):
        return len(self.texts)
    def __getitem__(self,idx):
        text=str(self.texts[idx])
        inputs=self.tokenizer.encode_plus(
            text,
            None,
            add_special_tokens=True,
            max_length=self.max_len,
            padding="max_length",
            truncation=True
            )
        resp={
            "ids":torch.tensor(inputs["input_ids"],dtype=torch.long),
            "mask":torch.tensor(inputs["attention_mask"],dtype=torch.long),
            "token_type_ids":torch.tensor(inputs["token_type_ids"],dtype=torch.long),
            "targets":torch.tensor(self.targets["idx"],dtype=torch.long)
            #multi--self.targets["idx"] is a vector
            # single--"targets":torch.tensor(self.targets["idx"],dtype=torch.float)
            }
        return resp
class TextModel(tez.model):
    def __init__(self,num_classes,num_train_steps):
        super().__init__
        self.bert=transformers.BertModel.from_pretrained(
            "bert-base-uncased",return_dict=False
        )
        self.bert_drop=nn.Dropout(0.3)
        self.out=nn.Linear(768,num_classes)
        self.num_train_steps=num_train_steps
        self.step_scheduler.after="batch"
    def fetch_optimizer(self):
        opt=AdamW(self.parameters(),lr=1e-4)
        return opt
    def fetch_scheduler(self):
        sch=get_linear_schedule_with_warmup(
            self.optimizer,num_warmup_steps=0,num_train_steps=self.num_train_steps
        )
        return sch
    def loss(self,outputs,targets):
        #single--return nn.BCEWithLogitsLoss()(outputs,targets.view(-1,1))
        return nn.CrossEntropyLoss()(outputs,targets)
    def monitor_metrics(self,outputs,targets):
        outputs=torch.argmax(outputs,axis=1).cpu().detach().numpy()>=0.5
       #single-- outputs=torch.sigmoid(outputs).cpu().detach().numpy()>=0.5
        targets=targets.cpu.detach().numpy()
        return{"accuracy": metrics.accuracy_score(targets,outputs)}
    def forward(self,ids,mask,token_type_ids,targets=None):
        _,x=self.bert(ids,attention_mask=mask,token_type_ids=token_type_ids)
        x=self.bert_drop(x)
        x=self.out(x)
        if targets is not None:
            loss=self.loss(x,targets)
            met=self.monitor_metrics(x, targets)
            return x,loss,met
        return x,0,{}
def train_model(fold):
    df=pd.read.csv("/home/datasets/imdb_folfs.csv")
    df_train=df[df.kfold !=fold].reset_index(drop=True)
    df_valid=df[df.kfold !=fold].reset_index(drop=True)
    
    train_dataset=Bertdataset(df_train.review.values, df_train.setiment.values)
    valid_dataset=Bertdataset(df_valid.review.values, df_valid.setiment.values)
    
    n_train_steps=int(len(df_train)/32*10)
    model=TextModel(num_classes=2, num_train_steps=n_train_steps)
   #single--class=1
   #single start comment:
    es=tez.callbacks.EarlyStopping(monitor="valid_loss",patience=3,model_path="model.bin")
    model.fit(train_dataset,valid_dataset=valid_dataset,device="cuda",epochs=10,train_bs=32,
              callbacks=[es])
    #single comment end
    model.load("model.bin",device="cuda")
    preds=model.predict(valid_dataset,device="cuda")
    for p in preds:
        print(p)
if  __name__=="__main__" :
    train_model(fold=0)