'''
Created on 2016年8月24日

@author: lil03
'''

# -*- coding=utf-8 -*-
from xml.etree.ElementTree import ElementTree,Element
import os,time
from Test_PubScripts.getExcelValue import getExcelValue


curPath = os.path.dirname(__file__)
XmlPath = curPath.replace('Test_PubScripts','Test_Data\DataRecord.xml')
class XmlAction():
    
    
    def __init__(self):
        self.ObjectList = getExcelValue().ObjectList
        self.InputList = getExcelValue().InputList
        self.CheckList = getExcelValue().CheckList
        self.GlobalList = getExcelValue().GlobalList
        self.xmlPath = XmlPath
    def initXmlfile(self,A,B,C,D):
        #初始化xml文件，
        #默认生成Excel转化xml的数据文件
        #未更新RunningData
        #A,B,C,D 表示对应的4个data字典
#         A = {'OAtton1':('id','123'),'OMtton2':('id','234','3'),'OWtton3':('id','345')}
#         B = {'DAlue1':'123','DMlue2':'234','DWlue3':'345'}
#         C = {'GA1':'1','GW2':'2','GM3':'3','hah':'22222'}
#         D = {'RA':'heihei','RW':'gaga','RA2':'xixixi'}
        BookName = {'ObjectData':A,
                    'InputData':B,
                    'CheckData':C,
                    'GlobalData':D,
                    'RunningData':{}
                    }
        Root = Element('OkDeerProject')
        Tree = ElementTree()
        Tree._setroot(Root)
        for kb in BookName:
            #ObjectData生成2层数据
            if kb == 'ObjectData':
                Dataroot = Element(kb)
                Root.append(Dataroot)
                for k in BookName[kb].keys():
                    v = BookName[kb][k]
                    if len(v) == 2:
                        adddict = {}
                        adddict[v[0]] = v[1]
                    elif len(v) == 3:
                        adddict = {}
                        adddict[v[0]] = v[1]
                        adddict['Index'] = v[2]
                    Dataroot.append(Element(k,adddict))
            else:
                DataRoot = Element(kb,BookName[kb])
                Root.append(DataRoot)
        Tree.write(self.xmlPath,xml_declaration=True, encoding='utf-8', method="xml")
        
    def getXmlTree(self,xmlPath = None):
        #解析Xml，返回解析结果树
        if xmlPath == None:
            xmlPath = self.xmlPath
        try:
            Tree = ElementTree()
            Tree.parse(xmlPath)
            return Tree
        except:
            print('Xml文件打开失败//')

    def setXmlData(self,name,value,Node = 'RunningData'):
        #添加xml运行数据
        Tree = self.getXmlTree()
        try:
            #存在，则直接修改
            ActiveNode = Tree.find(Node)
            ActiveNode.set(name,value)
            print('已修改运行数据<%s:%s>'%(name,value))
        except:
            #如果寻找对象出错，则代表对应节点不存在
            #则添加对应节点记录
            ObjectNode = Tree.find(Node.split(r'/')[0])
            NewNode = Element(Node.split(r'/')[1],{name:value})
            ObjectNode.append(NewNode)
            print('已新增运行数据<%s:%s>'%(name,value))
        Tree.write(self.xmlPath)

    def setXmlTreeByDict(self,Tree,args,Node = 'RunningData'):
        #直接添加dict数据
        for n in args:
            try:
                ActiveNode = Tree.find(Node)
                ActiveNode.set(n,args[n])
            except:
                #如果寻找对象出错，则代表对应节点不存在
                #则添加对应节点记录
                ObjectNode = Tree.find(Node.split(r'/')[0])
                NewNode = Element(Node.split(r'/')[1],{n:args[n]})
                ObjectNode.append(NewNode)
        return Tree
    
    def saveXml(self,Tree):
        Tree.write(self.xmlPath)

    def getXmlData(self,DataName,Node = 'RunningData'):
        #获取XmlData
        #DataName 属性名
        #Node：数据节点
        try:
            Tree = self.getXmlTree()
            ActiveNode = Tree.find(Node)
            DataValue = ActiveNode.get(DataName)
            return DataValue
        except:
            print('获取数据失败，请检查Xml节点：%s，变量名：%s'%(DataName,Node))

    def getTestData(self,Node = 'RunningData'):
        try:
            Tree = self.getXmlTree()
            ActiveNode = Tree.find(Node)
            return ActiveNode
        except:
            print('获取运行数据失败，请检查Xml节点：%s'%(ActiveNode))

    def initXmlDataFile(self):
        #初始化xml数据，加载解析Excel对应的对象以及用例数据
        XmlAction().initXmlfile(self.ObjectList,
                                self.InputList,
                                self.CheckList,
                                self.GlobalList)
        
    def upDateXmlData(self):
        #判断是否存在xml,不然则初始化xml
        if not os.path.exists(XmlPath):
            self.initXmlDataFile()
            print('------------------执行xml文件初始化----------------')
        else:
            print('Xml数据源同步中......')
        #更新除RunningData外的所有数据
        #更新InputData，checkData,Globallist
        NewTree = self.getXmlTree()
        NewTree = self.setXmlTreeByDict(NewTree,self.InputList,'InputData')
        NewTree = self.setXmlTreeByDict(NewTree,self.CheckList,'CheckData')
        NewTree = self.setXmlTreeByDict(NewTree,self.GlobalList,'GlobalData')
        #更新ObjectData
        ObjNodes = self.ObjectList.keys()
        for oj in ObjNodes:
            attrdict = {}
            attrName = self.ObjectList[oj][0]
            attrValue = self.ObjectList[oj][1]
            attrdict[attrName] = attrValue
            Node = r'ObjectData/' + oj
            NewTree = self.setXmlTreeByDict(NewTree,attrdict,Node)
            if len(self.ObjectList[oj]) == 3:
                indexdict = {}
                indexName = 'Index'
                indexValue = self.ObjectList[oj][2]
                indexdict[indexName] = indexValue
                Node = r'ObjectData/' + oj
                NewTree = self.setXmlTreeByDict(NewTree,indexdict,Node)
        self.saveXml(NewTree)
        print('Xml数据源同步已完成')

if __name__ == '__main__':
    datalist = XmlAction().getTestData('CheckData')
    pass