#! /usr/bin/env python
#coding=utf-8

import os
import cx_Oracle
import re
from dbUtils import *

CONNECT = "yxjc/Fd_yxjc_2ws@FDGXGZCX"
ENCODE = "gbk"
FILENAME = u"\\数据库文档.txt".encode(ENCODE)
FLAG = 1

tableList = []

#if FLAG == 1:
conn = cx_Oracle.connect(CONNECT) 
cursor = conn.cursor() 
cursor.execute("select table_name from user_tables order by table_name desc")
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
    cursor.execute("select t.TABLE_NAME, t.COLUMN_NAME,t.DATA_TYPE,t.DATA_LENGTH,t.NULLABLE,t.DATA_DEFAULT,c.COMMENTS from user_tab_columns t,user_col_comments c where t.table_name=c.TABLE_NAME and t.COLUMN_NAME=c.COLUMN_NAME and t.table_name = '%(tablename)s' order by column_name asc" % {"tablename":tableTmp.Name})
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

fileHandler = open(u"yxjc数据库文档.txt","w")
for table in tableList:
    tableLen = len(table.Records)
    fileHandler.write(table.Name+"\n")
    fileHandler.write(table.Remark+"\n")
    for i in range(tableLen):
        j = i+3
        record = table.Records[i]
        fileHandler.write(str(record.Name)+"\r")
        fileHandler.write(str(record.Type)+"\r")
        fileHandler.write(str(record.Length)+"\r")
        fileHandler.write(str(record.Nullable)+"\r")
        fileHandler.write(str(record.Alias)+"\r")
        fileHandler.write(str(record.Default)+"\r")
        fileHandler.write(str(record.Remark)+"\r")
        fileHandler.write("\r");
    fileHandler.write("\r");
fileHandler.write("\r");
            



