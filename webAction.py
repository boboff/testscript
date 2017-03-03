'''
Created on 2016年4月19日

@author: lil03
'''
# -*- coding: utf-8 -*-
from Test_PubScripts.getExcelValue import getExcelValue, filepath
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.select import Select
import datetime,sys,time,os
from time import sleep
from Test_PubScripts.dbAction import dbAction
import subprocess
from Test_PubScripts.XmlAction import XmlAction
#兼容本地执行以及远程执行。
#导入本地浏览器对象、以及远程设备remote对象
from selenium.webdriver import Remote
from selenium.webdriver.chrome.webdriver import WebDriver as Chrome
from selenium.webdriver.firefox.webdriver import WebDriver as Firefox

#导入浏览器配置类---#一个与浏览器相关的字典
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
# WebDriver = None
#######浏览器种类类型
## firefox,chrome,ie...
#设备列表
BrowserType = 'chrome'
# BrowserType = 'http://10.20.100.129:5555/wd/hub'
# importcontent = 'from selenium.webdriver.%s.webdriver import WebDriver'%BrowserType
# exec(importcontent)

curPath = os.path.dirname(__file__)

#创建一个获取对象的类
class webAction(Chrome,Firefox,Remote):
    #实例化初始化。
    def __init__(self):
        #{'baidufeild': ['id', 'kw']}
        self.ObjectList = getExcelValue().ObjectList       #对象列表（excel中的对象信息）
        #{'baidutext': '搜索一下Python'}
        self.InputList  = getExcelValue().InputList        #输入列表（excel中的输入信息）
        self.CheckList  = getExcelValue().CheckList        #检查列表（excel中的检查信息）
        self.GlobalList = getExcelValue().GlobalList  #全局变量
        if BrowserType == 'chrome':
            Chrome.__init__(self)
        elif BrowserType == 'firefox':
            Firefox.__init__(self)
        elif BrowserType:
            Remote.__init__(self,
                            command_executor = BrowserType,  
                            desired_capabilities=DesiredCapabilities.CHROME)

    def browserSetting(self):
        self.maximize_window()
        self.implicitly_wait(20)
        self.verificationErrors = []

    def GlobalValueAdd(self,*args):
        #添加全局变量
        #GlobalName：全局变量名
        #GlobalValue：变量值
        GlobalName = args[0]
        GlobalValue = ''
        if GlobalName in self.GlobalList.keys():
            print('%s已存在，已执行覆盖操作'%GlobalName)
        if len(args) == 2:
            GlobalValue = args[1]
        else:
            GlobalValue = self.getValueInput(GlobalName)
        self.GlobalList[GlobalName] = GlobalValue
        print('已添加全局变量<%s:%s>'%(GlobalName,GlobalValue))

    def getObjText_As_GlobalVlue(self,objectName):
        
        #添加objectName和对象Text值存为全局变量
        GlobalName = self.getValueInput(objectName)
        GlobalValue = self.getObject(objectName).text
        self.GlobalValueAdd(GlobalName, GlobalValue)

    def getValueGlobal(self,objectName):
        #获取全局变量
        #GlobalName：全局变量名
#         if objectName in self.GlobalList.keys():
#             return self.GlobalList[objectName]
#         else:
#             print('对象名称<%s>不存在，请检查GlobalData'%objectName)
        #返回全局变量数据，若不存在则返回自身
        objectName = self.GlobalList.get(objectName,objectName)
        return objectName

    def getValueInput(self,objectName):
        #获取Excel对象输入值内容  
#         if objectName in self.InputList.keys():
#             return self.InputList[objectName]
#         else:
#             print('对象名称<%s>不存在，请检查UnitCaseData'%objectName)
        #返回输入值数据，若不存在则返回自身
        objectName = self.InputList.get(objectName,objectName)
        return objectName
            
    def getValueCheck(self,objectName):     
        #获取Excel对象检查值
#         if objectName in self.CheckList.keys():
#             return self.CheckList[objectName]
#         else:
#             print('对象名称<%s>不存在，请检查UnitCaseData'%objectName)
        #返回检查值数据，若不存在则返回自身
        objectName = self.CheckList.get(objectName,objectName)
        return objectName
        
    def getValueObject(self,objectName):     
        #获取Excel对象属性值
#         if objectName in self.ObjectList.keys():
#             return self.ObjectList[objectName]
#         else:
#             print('对象名称<%s>不存在，请检查ObjectSource'%objectName)
        #返回检查值数据，若不存在则返回自身
        objectName = self.ObjectList.get(objectName,objectName)
        return objectName
    
#---------------------------------------------------------
        #组合属性类型以及对应属性值.定位         
        
    def getObjInfo(self,objProperty,objProValue):
        #单对象直接定位
        Element = self.find_element(objProperty.lower().replace('_',' '),objProValue)
        return Element
    
    def getObjsInfo(self,objProperty,objProValue):
        #定位对象集合
        Elements = self.find_elements(objProperty.lower().replace('_',' '),objProValue)
        return Elements
    
    def getSecondObjInfo(self,FirstObjPro,FirstObjValue,SecondObjPro,SecondObjValue):
        ##二层对象，层级定位
        Element = self.find_element(FirstObjPro.lower().replace('_',' '),FirstObjValue)\
        .find_element(SecondObjPro.lower().replace('_',' '),SecondObjValue)
        return Element
    
    def getSecondObjsInfo(self,FirstObjPro,FirstObjValue,SecondObjPro,SecondObjValue):
        ##二层对象，层级定位
        #一级对象唯一,二级对象为集合
        Elements = self.find_element(FirstObjPro.lower().replace('_',' '),FirstObjValue)\
        .find_elements(SecondObjPro.lower().replace('_',' '),SecondObjValue)
        return Elements
    
    
    def getObject(self,objectName,TopObjName = None):	    
        #获取Web一般控件对象
        #Ex：getObject(xx,xx)
        
        if isinstance(objectName, list) or isinstance(objectName, tuple):
            ObjectInfo = objectName
        else:
            ObjectInfo = self.getValueObject(objectName)
        #如果传入2个objectName，则使用两层定位。
        if TopObjName:
            if TopObjName in self.ObjectList.keys():
                TopObjInfo = self.getValueObject(TopObjName)
            else:
                TopObjInfo = TopObjName
            array = (TopObjInfo[0].lower(),TopObjInfo[1],ObjectInfo[0].lower(),ObjectInfo[1])
#             print(array)
            Element = self.getSecondObjInfo(*array)
        #不然则按照1个对象，定位对象x
#         如果存在序号，则获得集合，根据序号定位
        elif len(ObjectInfo) > 2:
            index = int(ObjectInfo[2])
            Element = self.getObjects(objectName)[index]
        else:
            Element = self.getObjInfo(ObjectInfo[0], ObjectInfo[1])
        return Element
    
    
    def getObjects(self,objectName,TopObjName = None):    
        #获取一个对象集
        #如果参数不在列表内，报错提示
        if isinstance(objectName, list) or isinstance(objectName, tuple):
            ObjectInfo = objectName
        else:
            ObjectInfo = self.getValueObject(objectName)
        if TopObjName:
            if TopObjName in self.ObjectList.keys():
                TopObjInfo = self.getValueObject(TopObjName)
            else:
                TopObjInfo = TopObjName
            array = (TopObjInfo[0].lower(),TopObjInfo[1],ObjectInfo[0].lower(),ObjectInfo[1])
            Element = self.getSecondObjsInfo(*array)
        else:
            Element = self.getObjsInfo(ObjectInfo[0], ObjectInfo[1])
        return Element
    
    def getObjectText(self,objectName): 
        #获取对象显示文本内容
        #TextName选填：log打印内容
        #Ex: xxx显示文本：yyyy
        ObjectText = self.getObject(objectName).text
        return ObjectText
    
    def browserUrl(self,DataName):   
        #进入对应链接
        if DataName in self.GlobalList.keys():
            url = self.getValueGlobal(DataName)
        elif DataName in self.InputList.keys():
            url = self.getValueInput(DataName)
        else:
            url = DataName
        self.get(url)

    def browserTabClose(self):         
        #关闭当前页面
        self.close()
        
    def browserQuit(self):          
        #关闭浏览器
        self.quit()

    def browserBack(self):          
        #浏览器回退
        self.back()

    def browserForward(self):       
        #浏览器前进
        self.forward()
        
    def browserRefresh(self):       
        #浏览器刷新
        self.refresh()

    def browsertoBottom(self):             
        #滚动滚动条至最底部
        Jscript = 'var q=document.documentElement.scrollTop=10000'
        self.execute_script(Jscript)

    def getWebShot(self,filename = None):            
        #截图
        #获取当前时间(定义的格式)
        date = datetime.datetime.now()
        nowdate = date.strftime('%Y-%m-%d')
        nowtime = date.strftime('%H-%M-%S')
        if filename == None:
            filename = nowtime
        #获取调用函数的路径
        filePath = sys.path[0]
        if 'TestCase' in filePath:
            PicShotPath = filePath.replace('TestCase',r'Test_log\ScreenShot\%s\%s.png'%(nowdate,filename))
        else:
            PicShotPath = filePath + r'\Test_log\ScreenShot\%s\%s.png'%(nowdate,filename)
        os.makedirs(os.path.dirname(PicShotPath),exist_ok=True)
        self.get_screenshot_as_file(PicShotPath)
#         print('已完成截图，图片存址：%s'%PicShotPath)
        return PicShotPath

    def objectClick(self,objectName,TopObjName = None):   
        #对象单击
        #支持2层定位，模块定位对象放置最后，或者通过k:v,入参
        #EX:
        #print(self.getValueObject(objectName))
        self.getObject(objectName,TopObjName).click()
        
    def objectClear(self,objectName,TopObjName = None):
        #输入框清除操作
        Element = self.getObject(objectName, TopObjName)
        Element.clear()
        
    def objectSet(self,objectName,DataName,TopObjName = None):     
        #对象输入(先清除在输入)
        #DataName:为CaseData(Excel)中的对象名
        #支持2层定位，模块定位对象放置最后，或者通过k:v,入参
        Element = self.getObject(objectName,TopObjName)
        if DataName in self.InputList.keys():
            sendInfo = self.getValueInput(DataName)
            Element.clear()
        #如果对象属于全局对象，则解析成全局变量
        elif DataName in self.GlobalList.keys():
            sendInfo = self.getValueGlobal(DataName)
            Element.clear()
        #如果对象是运行数据，则直接取xml中的值
        elif DataName[0:2] == 'RD' or DataName[0:4] == 'test':
            sendInfo = XmlAction().getXmlData(DataName)
            Element.clear()
        else:
            sendInfo = DataName
            try:
                Element.clear()
            except:
                pass
        Element.send_keys(sendInfo)
        
    def objectSet_Like_Mouse(self,objectName,DataName):
        Element = self.getObject(objectName)
        if DataName in self.InputList.keys():
            sendInfo = self.getValueInput(DataName)
            Element.clear()
        #如果对象属于全局对象，则解析成全局变量
        elif DataName in self.GlobalList.keys():
            sendInfo = self.getValueGlobal(DataName)
            Element.clear()
        #如果对象是运行数据，则直接取xml中的值
        elif DataName[0:2] == 'RD' or DataName[0:4] == 'test':
            sendInfo = XmlAction().getXmlData(DataName)
            Element.clear()
        else:
            sendInfo = DataName
            try:
                Element.clear()
            except:
                pass
        ActionChains(self).send_keys_to_element(Element,sendInfo).perform()
        
    def object_IsVisiable(self,objectName,TopObjName = None):
        #验证对象是否可见
        #支持2层定位，模块定位对象放置最后，或者通过k:v,入参
        Result = self.getObject(objectName,TopObjName).isvisible()
        return Result
        
    def objectDoubleClick(self,objectName,TopObjName = None):     
        #对象双击
        #支持2层定位，模块定位对象放置最后，或者通过k:v,入参
        element = self.getObject(objectName,TopObjName)
        ActionChains(self).double_click(element).perform()

    def objectRightClick(self,objectName,TopObjName = None):  
        #对象右击
        #支持2层定位，模块定位对象放置最后，或者通过k:v,入参
        element = self.getObject(objectName,TopObjName)
        ActionChains(self).context_click(element).perform()
        
    def objectMouseOver(self,objectName,TopObjName = None):   
        #鼠标滑至对象
        #支持2层定位，模块定位对象放置最后，或者通过k:v,入参
        element = self.getObject(objectName,TopObjName)
        ActionChains(self).move_to_element(element).perform()
        
    def objectsClick(self,objectName,DataName = None,TopObjName = None,index = 0):  
        #对象集点击,默认点击第一个
        if DataName:
            DB_Value = self.getValueInput(DataName)
            DB_Value_list = len(DB_Value).split(',')
            if  DB_Value_list == 1:
                clickIndex = DB_Value
            else:
                clickIndex = DB_Value_list[0]
        else:
            clickIndex = index
        self.getObjects(objectName,TopObjName)[clickIndex].click()
        
    def objectClick_Like_Mouse(self,objectName,TopObjName = None):    
        #模拟鼠标点击
        #支持2层定位，模块定位对象放置最后，或者通过k:v,入参
        element = self.getObject(objectName,TopObjName)
        ActionChains(self).click(element).perform()
        
    def objectSelect(self,objectName,DataName,TopObjName = None):  
        #对象选择(下拉框),通过显示值
        #支持2层定位，模块定位对象放置最后，或者通过k:v,入参
        element = self.getObject(objectName,TopObjName)
        if DataName in self.InputList.keys():
            selectValue = self.getValueInput(DataName)
        else:
            selectValue = DataName
        Select(element).select_by_value(selectValue)
        
    def objectExist(self,objectName,TopObjName = None,time = 3):   
        #判断对象是否存在
        try:
            self.implicitly_wait(time)
            self.getObject(objectName,TopObjName)
            IsExist = True
            self.implicitly_wait(20)
        except:
            IsExist = False
        return IsExist
        
    def objectIsdisplayed(self,objectName,TopObjName = None,time = 3):   
        #判断对象是否可操作
        self.implicitly_wait(time)
        Isdisplayed = self.getObject(objectName, TopObjName).is_displayed()
        self.implicitly_wait(20)
        return Isdisplayed
        
    def WaitObjAppear(self,objectName,WaitTime = 20,TopObjName = None):   
        #等待对象出现(默认最大等待时间20s)
        ActionTime = time.time()
        IsExist = False
        while True:
            try:
                self.getObject(objectName,TopObjName)
                IsExist = True
            except:
                pass
            NewTime = time.time()
            if IsExist == True:
                #时间保留2位小数
#                 print('已等待出现:<%s>,大约等待(%.2f)s'%(objectName,round((NewTime-ActionTime),2)))
                break
            elif NewTime >= ActionTime+WaitTime:
                print('等待超时:<%s>,大约等待(%.2f)s'%(objectName,WaitTime))
                break
            else:
                time.sleep(0.5)
        
    def WaitObjDisAppear(self,objectName,TopObjName = None,WaitTime = 30):
        #等待对象消失(默认最大等待时间20s)
        ActionTime = time.time()
        IsExist = True
        while True:
            try:
                self.getObject(objectName,TopObjName)
            except:
                IsExist = False
            NewTime = time.time()
            if IsExist == False:
                #时间保留2位小数
                print('已等待消失:<%s>,大约等待(%.2f)s'%(objectName,round((NewTime-ActionTime),2)))
                break
            elif NewTime >= ActionTime+WaitTime:
                print('等待超时:<%s>,大约等待(%.2f)s'%(objectName,WaitTime))
                break
            else:
                time.sleep(0.5)
        
    def objTextCheck(self,objectName,DataName,TopObjName = None):  
        #检查文本内容是否符合预期
        CheckObject = self.getObject(objectName,TopObjName)
        ObjectText = CheckObject.text
        if DataName in self.CheckList.keys():
            CheckText = self.getValueCheck(DataName)
        #如果对象属于全局对象，则解析成全局变量
        elif DataName in self.GlobalList.keys():
            CheckText = self.getValueGlobal(DataName)
        else:
            CheckText = DataName

        if ObjectText != CheckText:
            print('验证:<%s>,预期结果"%s",实际结果"%s"'%(objectName,CheckText,ObjectText))
        assert CheckText == ObjectText

    def iframeTo(self,objectName):
        #切换iframe
        Object = self.getObject(objectName)
        self.switch_to.frame(Object)
        time.sleep(1)
        
    def iframeOut(self):                
        #切出iframe
        self.switch_to.default_content()
        
    def windowsTo(self,Index):
        #Index代表序号，即操作第几个Page页
        Allhandle = self.window_handles
        self.switch_to.window(Allhandle[Index-1])
        
    def AnnexUpload(self,AnnexName = '附件测试001.png'):
        #附件上传
        #AnnexName:附件目录下的文件名称(需要带上图片格式)
        #附件目录：项目下的\Test_Resource\Annex
        #exePath，上传附件exe的地址。
        
        #如果执行机为本机，则运行相对路径下Exe上传附件程序
        if BrowserType in ('chrome',
                           'firefox'):
            #获取调用函数的路径
            exePath = curPath.replace('Test_PubScripts',r'Test_Data\AnnexExe\AnnexUpload.exe')
#             filePath = sys.path[0]
            filePath = os.path.abspath(os.curdir) #获取当前所执行程序所在的绝对路径
#             print('filePath的值为：'+filePath)            

            if 'TestCase' in filePath:
                AnnexFilePath = filePath.replace('TestCase',r'Test_Resource\Annex')

            else:
                AnnexFilePath = filePath + r'\Test_Resource\Annex'

            AnnexPath = AnnexFilePath + '\\' + AnnexName
#             print('需上传的图片路径:'+ AnnexPath)
            #拼接执行shell
            command = '"'+exePath+'"'+' '+'"'+BrowserType+'"'+' '+'"'+AnnexPath+'"'
        else:
            #如果为远程调用，默认调用桌面附件路径
            exePath = r'C:\Users\Administrator\Desktop\SeleniumAnnex\AnnexUpload.exe'
            AnnexPath = r'C:\Users\Administrator\Desktop\SeleniumAnnex' + '\\' + AnnexName
            command = exePath + ' ' + '"' + 'chrome' + '"' + ' ' + '"' + AnnexPath + '"'
#         print(command)
        Action = subprocess.Popen(command,shell = True)
        Action.wait(10)
        Action.kill()
        sleep(1)

    def JSAction(self,Jscript):
        #操作JS
        self.execute_script(Jscript)
        
    def objectClick_Bytext(self,TopObjName,DB_TopMenu):
        ##根据模块下的对象标签中text的值来定位并且点击对象
        #TopObjName 模块对象
        #TagType为寻找对象的tag类型
        #遍历模块对象下满足text条件的对象并点击
        ExcelValue = self.getValueInput(DB_TopMenu)
        Textcontent = ExcelValue.split(',')[0]
        TagType = ExcelValue.split(',')[-1]
        ObjectList = self.getObject(TopObjName).find_elements_by_tag_name(TagType)
        for everyobj in ObjectList:
            #print(everyobj.text + 'Ok')
            if everyobj.text == Textcontent:
                everyobj.click()
#                 print('已完成点击:<%s>'%(TopObjName))
                break

    def get_sqlResultDict(self,SQLName,user,passwd,host,db):
        #获取数据库结果数据
        #链接数据库相关
        #解析Excel传入变量名
        SqlContent = self.getValueInput(SQLName)
        user = self.getValueGlobal(user)
        passwd = self.getValueGlobal(passwd)
        host = self.getValueGlobal(host)
        db = self.getValueGlobal(db)
        #创建dbAciton类，获取返回结果
        dbconn = dbAction(user,passwd,host,db)
        return dbconn.getSqlResultDict(SqlContent)
        
    def get_SqlResultCount(self,SQLName,user,passwd,host,db):
        #获取数据结果行数
        SqlContent = self.getValueInput(SQLName)
        user = self.getValueGlobal(user)
        passwd = self.getValueGlobal(passwd)
        host = self.getValueGlobal(host)
        db = self.getValueGlobal(db)
        #创建dbAciton类，获取返回结果
        dbconn = dbAction(user,passwd,host,db)
        return dbconn.getSqlResultCount(SqlContent)
    
    def logPage(self,Text):
        #页面表示log日志
        strLen = len(str(Text))
        MaxLen = 60
        if strLen%2 == 0:
            count = int((MaxLen-strLen*2)/2)
            print('='*count+Text+'='*count)
        else:
            count = int((MaxLen-strLen*2)/2)
            print('='*(count-1)+ Text + '='*count)
    
if __name__ == '__main__':
    pass