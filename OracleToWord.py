#! /usr/bin/env python
#coding=utf-8

import os
import cx_Oracle
import re
from dbUtils import *
from win32com import client as wc
import pywintypes

CONNECT = "yxjc/Fd_yxjc_2ws@FDGXGZCX"
ENCODE = "gbk"
FILENAME = u"\\数据库文档.doc".encode(ENCODE)
FLAG = 1

tableList = []

#if FLAG == 1:
conn = cx_Oracle.connect(CONNECT) 
cursor = conn.cursor() 
cursor.execute("select table_name from user_tables")
tables = cursor.fetchall()
for table in tables:
    tableTmp = DBTable()
    tableTmp.Name = table[0]
    cursor.execute("select   comments   from   user_tab_comments where table_name = '%(tablename)s'" % {"tablename":tableTmp.Name})
    temp = cursor.fetchone()[0]
    if temp != None:
        tableTmp.Remark = temp
    cursor.execute("select b.column_name  from USER_CONSTRAINTS a, USER_CONS_COLUMNS b where a.constraint_name = b.constraint_name and a.constraint_type ='P' and a.table_name = '%(tablename)s'"%{"tablename":tableTmp.Name})
    keys = cursor.fetchall()
    for key in keys:
        tableTmp.Keys = tableTmp.Keys + key[0] +","
    if tableTmp != "":
        tableTmp.Keys = tableTmp.Keys[0:-1]
    cursor.execute("select t.TABLE_NAME, t.COLUMN_NAME,t.DATA_TYPE,t.DATA_LENGTH,t.NULLABLE,t.DATA_DEFAULT,c.COMMENTS from user_tab_columns t,user_col_comments c where t.table_name=c.TABLE_NAME and t.COLUMN_NAME=c.COLUMN_NAME and t.table_name = '%(tablename)s'" % {"tablename":tableTmp.Name})
    cols = cursor.fetchall()
    for col in cols:
        record = DBTableRecord()
        record.Name = col[1]
        record.Type = col[2]
        record.Length = "("+str(col[3])+")"
        if col[4] == "N":
            record.Nullable = False
        if col[5] != None:
            record.Default = str(col[5])
        if col[6] != None:
            record.Alias = col[6]
        tableTmp.Records.append(record)
    tableList.append(tableTmp)
conn.close()

if FLAG == 1:
    word = wc.Dispatch("Word.Application")
    userOvertype = word.Options.Overtype 
    if userOvertype == True:
        word.Options.Overtype = False
    dir = os.getcwd()
    filePath = dir+FILENAME
    print filePath
    e = os.path.exists(filePath)
    if e == True:
        os.unlink(filePath)
    e = os.path.exists(filePath)
    if e != True:
        doc = word.Documents.Add()
        doc.SaveAs(filePath)

    doc = word.Documents.Open(filePath)
    sel = word.Selection #获取selection对象
    word.Options.CheckSpellingAsYouType = False
    word.Options.CheckGrammarAsYouType = False
    
    myRange = doc.Range(0,0) #定义光标的位置
    myRange.InsertAfter(u"数据库文档\n") 
   
    for table in tableList:
        tableLen = len(table.Records)
        #移动光标
        myRange = doc.Range(myRange.End,myRange.End)
        tableTitle = "\n"+table.Remark+" "+table.Name+"\n"
        myRange.InsertBefore(tableTitle)
        #移动光标
        myRange = doc.Range(myRange.End,myRange.End)
        wordtable = doc.Tables.Add(myRange ,3+tableLen,7)
        wordtable.Style = u"网格型"   
        #重新分配表结构
        tableRange = doc.Range(wordtable.Rows[0].Cells[1].Range.Start,wordtable.Rows[0].Cells[4].Range.End)
        tableRange.Cells.Merge()
        tableRange = doc.Range(wordtable.Rows[0].Cells[2].Range.Start,wordtable.Rows[0].Cells[3].Range.End)
        tableRange.Cells.Merge()
        tableRange = doc.Range(wordtable.Rows[1].Cells[1].Range.Start,wordtable.Rows[1].Cells[6].Range.End)
        tableRange.Cells.Merge()
        #写表头
        wordtable.Rows[0].Cells[0].Range.Text=u"表名".encode(ENCODE)
        wordtable.Rows[0].Cells[1].Range.Text=table.Name
        wordtable.Rows[0].Cells[2].Range.Text=table.Remark   
        wordtable.Rows[1].Cells[0].Range.Text=u"主键".encode(ENCODE)
        wordtable.Rows[1].Cells[1].Range.Text=table.Keys  
        wordtable.Rows[2].Cells[0].Range.Text=u"字段名称".encode(ENCODE)
        wordtable.Rows[2].Cells[1].Range.Text=u"类型".encode(ENCODE)
        wordtable.Rows[2].Cells[2].Range.Text=u"数据类型".encode(ENCODE)
        wordtable.Rows[2].Cells[3].Range.Text=u"是否可空".encode(ENCODE)
        wordtable.Rows[2].Cells[4].Range.Text=u"字段说明".encode(ENCODE)
        wordtable.Rows[2].Cells[5].Range.Text=u"缺省值".encode(ENCODE)
        wordtable.Rows[2].Cells[6].Range.Text=u"备注".encode(ENCODE)
        #写字段名称
        for i in range(tableLen):
            j = i+3
            record = table.Records[i]
            wordtable.Rows[j].Cells[0].Range.Text=record.Name
            wordtable.Rows[j].Cells[1].Range.Text=record.Type
            wordtable.Rows[j].Cells[2].Range.Text=record.Length
            wordtable.Rows[j].Cells[3].Range.Text=record.Nullable
            wordtable.Rows[j].Cells[4].Range.Text=record.Alias
            wordtable.Rows[j].Cells[5].Range.Text=record.Default
            wordtable.Rows[j].Cells[6].Range.Text=record.Remark            
        myRange = doc.Range(wordtable.Range.End, wordtable.Range.End)
        sel.MoveDown(4,1)  
    doc.Close(-1)



