'''
Created on 2016年7月4日

@author: lil03
'''
import xlrd,datetime
#获取当前时间并且格式化
date = datetime.datetime.now()
datevalue = date.strftime('%m%d%H%M')
#2天后的时间
AfterDate = date + datetime.timedelta(days = 2)


import os,re
curpath = os.path.dirname(__file__)
Path = curpath.replace('Test_PubScripts','Test_Data')
DirList = os.listdir(path=Path)
ObjDataList = []
CaseDataList = []
GlobalDataList = []
for dir in DirList:
    dirpath = None
    excelPath = os.path.join(Path,dir)
    if os.path.isfile(excelPath):
        #如果存在object或者unitcase 的Excel
        if 'object' in dir.lower()[:6]:
            if '.xlsx' in dir.lower():
            #获取objectSuorce 对象数据文件集合。
                ObjDataList.append(excelPath)
        if 'unitcase' in dir.lower()[:8]:
            #获取unitCaseTada 用例数据文件集合
            if '.xlsx' in dir.lower():
                CaseDataList.append(excelPath)
        if 'global' in dir.lower()[:6]:
            #获取GlobalData 用例数据文件集合
            if '.xlsx' in dir.lower():
                GlobalDataList.append(excelPath)
    if os.path.isdir(os.path.join(Path,dir)):
        #如果是文件夹
        dirpath = os.path.join(Path,dir)
        fileList = os.listdir(path=dirpath)
        for file in fileList:
            filepath = os.path.join(dirpath,file)
            if 'object' in file.lower()[:6]:
                #获取objectSuorce 对象数据文件集合。
                ObjDataList.append(filepath)
            if 'unitcase' in file.lower()[:8]:
                #获取unitCaseTada 用例数据文件集合
                CaseDataList.append(filepath)
            if 'global' in dir.lower():
                #获取GlobalData 用例数据文件集合
                GlobalDataList.append(excelPath)

class getExcelValue():

    def __init__(self):
        #对象信息标题列表
        self.ObjectPath = ObjDataList
        self.DataInfo = CaseDataList
        self.GlobalExcel = GlobalDataList
        self.GlobalList = self.getGlobalList()
        self.ObjectList  = self.getObjectList()
        self.InputList   = self.getInputList()
        self.CheckList   = self.getCheckList()
        #替换列表数据
        self.replaceInputList()
        self.replaceObjectList()
        self.replaceCheckList()

    def replaceObjectList(self):
        #替换对象列表
        ####替换对象列表数据
        for key,Data in self.ObjectList.items():
            #如果数据中含有+time，则以+分割str. 替换time为datevalue
            #ex:   活动店铺05231737 , 活动店铺+time
            if '+' in Data:
                Valuelist = Data.split('+')
                if 'time' in Valuelist:
                    Valuelist[Valuelist.index('time')] = datevalue
                    Data = ''.join(Valuelist)
#                             print('%s:%s'%(objectName,DataValue))
                self.ObjectList[key] = Data
        #若数据存在E$xxxx，则替换成InputList对应对象值。
        for Data in self.ObjectList.values():
            TextData = Data[1]
            #正则匹配，提取G$与$之间的替换字段
            #处理G$
            matchGobj = re.findall('(?<=G\{).*(?=\})',TextData)
            if len(matchGobj) != 0:
                MatchName = matchGobj[0]
                Changeword = r'G{'+ MatchName + r'}'
                TextData = TextData.replace(Changeword,self.GlobalList[MatchName])
                Data[1] = TextData
                
        for Data in self.ObjectList.values():
            TextData = Data[1]
            matchEobj = re.findall('(?<=E\{).*(?=\})',TextData)
            if len(matchEobj) != 0:
                for obj in matchEobj:
                    MatchName = obj
                    Changeword = r'E{'+ MatchName + r'}'
                    TextData = TextData.replace(Changeword,self.InputList[MatchName])
                    Data[1] = TextData

                            
    def replaceInputList(self):
        for key,Data in self.InputList.items():
            #如果数据中含有+time，则以+分割str. 替换time为datevalue
            #ex:   活动店铺05231737 , 活动店铺+time
            if '+' in Data:
                Valuelist = Data.split('+')
                if 'time' in Valuelist:
                    Valuelist[Valuelist.index('time')] = datevalue
                    Data = ''.join(Valuelist)
#                             print('%s:%s'%(objectName,DataValue))
                self.InputList[key] = Data
        #若数据存在E$xxxx，则替换成对应对象值。
        for key,Data in self.InputList.items():
            TextData = Data
            #正则匹配，提取G$与$之间的替换字段
            #处理G$
            matchGobj = re.findall('(?<=G\{).*(?=\})',TextData)
            if len(matchGobj) != 0:
                MathchName = matchGobj[0]
                Changeword = r'G{'+ MathchName + r'}'
                TextData = TextData.replace(Changeword,self.GlobalList[MathchName])
                self.InputList[key] = TextData
            
            matchEobj = re.findall('(?<=E\{).*(?=\})',TextData)
            if len(matchEobj) != 0:
                MathchName = matchEobj[0]
                Changeword = r'E{'+ MathchName + r'}'
                TextData = TextData.replace(Changeword,self.InputList[MathchName])
                self.InputList[key] = TextData


    def replaceCheckList(self):
        for key,Data in self.CheckList.items():
            #如果数据中含有+time，则以+分割str. 替换time为datevalue
            #ex:   活动店铺05231737 , 活动店铺+time
            if '+' in Data:
                Valuelist = Data.split('+')
                if 'time' in Valuelist:
                    Valuelist[Valuelist.index('time')] = datevalue
                    Data = ''.join(Valuelist)
#                             print('%s:%s'%(objectName,DataValue))
                self.InputList[key] = Data
        #若数据存在E$xxxx，则替换成对应对象值。
        for key,Data in self.CheckList.items():
            TextData = Data
            #正则匹配，提取G$与$之间的替换字段
            #处理G$
            matchGobj = re.findall('(?<=G\{).*(?=\})',TextData)
            if len(matchGobj) != 0:
                MathchName = matchGobj[0]
                Changeword = r'G{'+ MathchName + r'}'
                TextData = TextData.replace(Changeword,self.GlobalList[MathchName])
                self.CheckList[key] = TextData
                
            matchEobj = re.findall('(?<=E\{).*(?=\})',TextData)
            if len(matchEobj) != 0:
                MathchName = matchEobj[0]
                Changeword = r'E{'+ MathchName + r'}'
                TextData = TextData.replace(Changeword,self.InputList[MathchName])
                self.CheckList[key] = TextData



    def getObjectList(self):
        objectList = {}
        for Source in self.ObjectPath:
            FileList = self.getObjectList_ByExcel(Source)
            objectList = dict(objectList, **FileList)
        return objectList
            
    def getObjectList_ByExcel(self,ExcelPath):
        #获取ObjectSource.xlsx下对象信息
        objectList = {}
        try:
            Workbook = xlrd.open_workbook(ExcelPath)
        except:
            print('数据初始化异常，请检查Object对应Excel文件')
        #遍历sheet页的数据
        DataAssert = True
        for sheet in Workbook.sheet_names():
            curSheet = Workbook.sheet_by_name(sheet)
            #总行数
            rowCount = curSheet.nrows
            #第一列为标题，忽略掉。
            for row in range(1,rowCount):
                #1,2,3代表需求字段在Excel的列Index
                #1：对象名称
                #2：对象属性名
                #3：对象属性值
                objProperty =   str(curSheet.cell_value(row,2))
                ObjProValue =   str(curSheet.cell_value(row,3))
                objectName  =   str(curSheet.cell_value(row,1))
                try:
                    objIndex    =   str(curSheet.cell_value(row,4))
#                     print('无序号列')
                except:
                    pass
                #当前列存在重复则提示
                if objectName != '' and objProperty!= '':
                #属性以及属性值组成数组
                    PropertyArr = [objProperty,ObjProValue]
                    if objectName in objectList.keys():
                        print('当前对象名称%s存在重复，在<%s>页第%s行！'%(objectName,sheet,(row+1)))
                        DataAssert = False
                    #如果输入了序号项，则组成3元素列表。
                    if objIndex != '':
                        PropertyArr = [objProperty,ObjProValue,objIndex]
                #属性以及对象名称组成字典
                    objectList[objectName] = PropertyArr
        assert DataAssert
        return objectList
    
    def getInputList(self):
        InputList = {}
        for Source in self.DataInfo:
            FileList = self.getInputList_byExcel(Source)
            InputList = dict(InputList, **FileList)
        return InputList
            
    def getInputList_byExcel(self,DataPath):
        #获取UnitCaseData下数据
        try:
            Workbook = xlrd.open_workbook(DataPath)
        except:
            print('数据初始化异常，请检查Data对应Excel文件')
        InputList = {}
        #删选出Page_开头的Sheet页
        PageSheets = [x for x in Workbook.sheet_names() if 'Page_' in x]
        for sheet in PageSheets:
            #选择Page页对象加载为InputList
            curSheet = Workbook.sheet_by_name(sheet)
            #总行数
            rowCount = curSheet.nrows
            #取单元格值
            #第一列为标题，忽略掉。
            DataAssert = True
            for row in range(1,rowCount):
                objectName = str(curSheet.cell_value(row,2))
                DataValue  = str(curSheet.cell_value(row,3))
                if objectName != '' and DataValue != '':
                    #当前列存在重复则提示
                    if objectName in InputList.keys():
                        print('当前对象名称%s存在重复，在<%s>页第%s行！'%(objectName,sheet,(row+1)))
                        DataAssert = False
                    InputList[objectName] = DataValue
        assert DataAssert
        return InputList

    def getCheckList(self):
        CheckList = {}
        for Source in self.DataInfo:
            FileList = self.getCheckList_ByExcel(Source)
            CheckList = dict(CheckList, **FileList)
        return CheckList

    def getCheckList_ByExcel(self,DataPath):
        #获取检查值
        try:
            Workbook = xlrd.open_workbook(DataPath)
        except:
            print('数据初始化异常，请检查Data对应Excel文件')
        CheckList = {}
        PageSheets = [x for x in Workbook.sheet_names() if 'Page_' in x]
        for sheet in PageSheets:    
            curSheet = Workbook.sheet_by_name(sheet)
            #总行数
            rowCount = curSheet.nrows
            #总列数
            #colCount = curSheet.ncols
            #取单元格值
            #第一列为标题，忽略掉。
            DataAssert = True
            for row in range(1,rowCount):
                objectName = str(curSheet.cell_value(row,2))
                ChcekValue = str(curSheet.cell_value(row,4))
                #当前列存在重复则提示
                if objectName in CheckList.keys():
                    print('当前对象名称%s存在重复，在<%s>页第%s行！'%(objectName,sheet,(row+1)))
                    DataAssert = False
                #判断不为空
                elif ChcekValue != '' and objectName != '':
                    CheckList[objectName] = ChcekValue
        assert DataAssert
        return CheckList
    
    
    def getGlobalList(self):
        GlobalList = {}
        for DataExcel in self.GlobalExcel:
            FileList = self.getGlobalList_ByExcel(DataExcel)
            GlobalList = dict(GlobalList, **FileList)
        for key,Data in GlobalList.items():
            #如果数据中含有+time，则以+分割str. 替换time为datevalue
            #ex:   活动店铺05231737 , 活动店铺+time
            if '+' in Data:
                Valuelist = Data.split('+')
                if 'time' in Valuelist:
                    Valuelist[Valuelist.index('time')] = datevalue
                    Data = ''.join(Valuelist)
#                             print('%s:%s'%(objectName,DataValue))
                GlobalList[key] = Data
#         for k,v in GlobalList.items():
#             print(k,v)
        return GlobalList
    
    def getGlobalList_ByExcel(self,DataExcel):
        #获取全局
        Workbook = xlrd.open_workbook(DataExcel)
        PageSheets = [x for x in Workbook.sheet_names() if 'global' in x.lower()]
        GlobalVariableList = {}
        #如果有重复数据则停止脚本运行
        DataAssert = True
        for sheet in PageSheets:
            curSheet = Workbook.sheet_by_name(sheet)
            #总行数
            rowCount = curSheet.nrows
            for row in range(1,rowCount):
                GlobalName = str(curSheet.cell_value(row,1))
                GlobalValue = str(curSheet.cell_value(row,2))
                if GlobalName in GlobalVariableList.keys():
                    print('当前对象名称%s存在重复，在<%s>页第%s行！'%(GlobalName,sheet,row))
                    break
                    DataAssert = False
                elif GlobalName != '' and GlobalValue != '':
                    GlobalVariableList[GlobalName] = GlobalValue
        assert DataAssert
        return GlobalVariableList

if __name__ == '__main__':
    value = getExcelValue().CheckList['DMJXFWGLS_Result']
    print(value)