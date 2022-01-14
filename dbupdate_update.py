# -*- coding: utf-8 -*-
"""
Created on Tue Oct 12 13:08:15 2021

@author: ECOCNA_dev
"""

import mariadb


class DBCONN:
    
    def __intif__(self):
        super().__init__()
        self.Host
        self.Port
        self.DataBase
        self.User
        self.PassWord


    def SetDataBase(self,host,port,database,user,password):
        self.Host = host
        self.Port = port
        self.DataBase = database
        self.User = user
        self.PassWord = password


    def DB_CONN(self):
        """DB연결"""

        '''
        #DB config
        host = '127.0.0.1'
        port = int(3306)
        database = 'ieps_test2'
        user = 'root'
        passwd = '86rhddl0921'
        '''
        autocommit = False
        
        self.conn = mariadb.connect(user = self.User,password = self.PassWord,host = self.Host,port = int(self.Port),database = self.DataBase, autocommit = autocommit) #connection

        self.cur= self.conn.cursor()

        return "success"
  

    def query(self,table_name,fields):
        """쿼리"""
        sql_query = "SELECT " + fields +" FROM " + table_name  #query
        self.cur.execute(sql_query) #쿼리 실행
        resultset = self.cur.fetchall() # 
     
    
        return resultset


    def query_like(self,table_name,fields,like):
        """쿼리"""
        sql_query = "SELECT " + fields +" FROM " + table_name + " WHERE processNo LIKE " + "'"+ like + "-%' " #query
        self.cur.execute(sql_query)#쿼리 실행
        resultset = self.cur.fetchall() # 
    
        return resultset
