#! /usr/bin/env python
#coding=utf-8

class DBTable:
    def __init__(self):
        self.Name = ""
        self.Keys = ""
        self.Remark = ""
        self.Records = []
        
class DBTableRecord:
    def __init__(self):
        self.Name = ""
        self.Type = ""
        self.Length = ""
        self.Nullable = True
        self.Remark = ""
        self.Default = ""
        self.Alias = ""
