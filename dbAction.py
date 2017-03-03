'''
Created on 2016年7月26日

@author: lil03
'''
# -*- coding: utf-8 -*-

import pymysql

class dbAction():
    def __init__(self,
                 user = None,
                 passwd = None,
                 host = None,
                 db = None):
        #链接数据库

        self.user = user
        self.passwd = passwd
        self.host = host
        self.db = db

    def connectDB(self):
        conn = pymysql.connect(user = self.user,
                               passwd = self.passwd,
                               host = self.host,
                               db = self.db,
                               charset='utf8')
        return conn
        
    def getSqlResultDict(self,SqlContent):
        conn = self.connectDB()
        curConn = conn.cursor()
        #执行Sql，获取结果
        curConn.execute(SqlContent)
        #返回Sql结果
        #字典形式
        #demo：
        #{'id':[1,2,3,4,5]}
        ResultDict = {}
        #获取所有查询结果标题
        DictKey = []
        for desc in curConn.description:
            DictKey.append(desc[0])
        #根据标题个数，分别获取标题所有结果并加入字典value。
        tvlist  = []
        for i in range(len(DictKey)):
            for tv in curConn._rows:
                tvlist.append(tv[i])
            ResultDict[DictKey[i]] = tvlist
            tvlist = []
        if curConn.rowcount == 0:
            print('Sql查询无结果')
            #第一列加载完清空,继续第二列。
        #加载完关闭链接
        curConn.close()
        conn.close()
        #返回加载数据
        return ResultDict


    def getSqlResultCount(self,SqlContent):
        #获取查询结果行数
        conn = self.connectDB()
        curConn = conn.cursor()
        #执行Sql，获取结果
        curConn.execute(SqlContent)    
        return curConn.rowcount
        curConn.close()
        conn.close()
        
# dbconnect = dbAction('yschome',
#                      'test@2016',
#                      '10.20.102.51',
#                      'yscca')
# storeName = '唐慧荣的服务店1125'
# result = dbconnect.getSqlResultDict("select login_name from sys_user where user_name = '{}'".format(storeName))
# print(result)
# sellerStoreName = result['login_name'][0]