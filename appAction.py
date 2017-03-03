'''
Created on 2016年8月10日

@author: lil03
'''
from Test_PubScripts.getExcelValue import getExcelValue
import datetime,sys,os
import time
from Test_PubScripts.dbAction import dbAction
#父类由于需要确定到浏览器种类，所以动态加载
#######浏览器种类类型
## firefox,chrome,ie...
from appium import webdriver
from Test_PubScripts.XmlAction import XmlAction
from time import sleep

#创建一个获取对象的类
class appAction(webdriver.Remote):
    #实例化初始化。
    def __init__(self,host = 'http://localhost:4723/wd/hub'):
        self.desired_caps = {'platformName':'Android',
                             'platformVersion':'4.4.2',
                             'deviceName':'Android Emulator',
                             'appPackage':'com.trisun.vicinity.activity',
                             'appActivity':'com.trisun.vicinity.init.activity.SplashActivity',
                             'unicodeKeyboard':True,
                             'resetKeyboard':True}
        webdriver.Remote.__init__(self,host,self.desired_caps)
        #数据导入初始化
        #{'baidufeild': ['id', 'kw']}
        self.ObjectList = getExcelValue().ObjectList       #对象列表（excel中的对象信息）
        #{'baidutext': '搜索一下Python'}
        self.InputList  = getExcelValue().InputList        #输入列表（excel中的输入信息）
        self.CheckList  = getExcelValue().CheckList        #检查列表（excel中的检查信息）
        self.GlobalList = getExcelValue().GlobalList  #全局变量

    def close_Session(self):
        self.quit()

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
    
    def getObject(self,objectName):        
        #获取Web一般控件对象
        #Ex：getObject(xx,xx)
        if isinstance(objectName,list) or isinstance(objectName,tuple):
            ObjectInfo = objectName
        else:
            ObjectInfo = self.getValueObject(objectName)
#         如果存在序号，则获得集合，根据序号定位
        if len(ObjectInfo) > 2:
            index = int(ObjectInfo[2])
            Element = self.getObjects(objectName)[index]
        else:
            Element = self.getObjInfo(ObjectInfo[0], ObjectInfo[1])
        return Element
    
    def getObjects(self,objectName):    
        #获取一个对象集
        #如果参数不在列表内，报错提示
        if isinstance(objectName, list) or isinstance(objectName, tuple):
            ObjectInfo = objectName
        else:
            ObjectInfo = self.getValueObject(objectName)
        Element = self.getObjsInfo(ObjectInfo[0], ObjectInfo[1])
        return Element

    def getObjectText(self,objectName,TextName = '显示文本'): 
        #获取对象显示文本内容
        #TextName选填：log打印内容
        #Ex: xxx显示文本：yyyy
        ObjectText = self.getObject(objectName).get_attribute('text')
#         print('<%s>%s：%s'%(objectName,TextName,ObjectText))
        return ObjectText

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

    def objectClick(self,objectName):   
        #对象单击
        #支持2层定位，模块定位对象放置最后，或者通过k:v,入参
        #EX:
        self.getObject(objectName).click()
        
    def objectSet(self,objectName,DataName):     
        #对象输入(先清除在输入)
        #DataName:为CaseData(Excel)中的对象名
        #支持2层定位，模块定位对象放置最后，或者通过k:v,入参
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
        Element.send_keys(sendInfo)
        
    def object_IsVisiable(self,objectName):
        #验证对象是否可见
        #支持2层定位，模块定位对象放置最后，或者通过k:v,入参
        Result = self.getObject(objectName).isvisible()
        return Result
        
    def objectsClick(self,objectName,DataName = None):  
        #对象集点击,默认点击第一个
        clickIndex = self.getValueObject(objectName)[2]
        self.getObjects(objectName)[clickIndex].click()
        
    def objectExist(self,objectName,time = 3):   
        #判断对象是否存在，print结果
        try:
            self.implicitly_wait(time)
            self.getObject(objectName)
            IsExist = True
        except:
            IsExist = False
        return IsExist
        
    def WaitObjAppear(self,objectName,WaitTime = 30):   
        #等待对象出现(默认最大等待时间20s)
        ActionTime = time.time()
        IsExist = False
        while True:
            try:
                self.getObject(objectName)
                IsExist = True
            except:
                IsExist = False
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
        
    def WaitObjDisAppear(self,objectName,WaitTime = 30):
        #等待对象消失(默认最大等待时间20s)
        ActionTime = time.time()
        IsExist = True
        while True:
            try:
                self.getObject(objectName)
                IsExist = True
            except:
                IsExist = False
            NewTime = time.time()
            if IsExist == False:
                #时间保留2位小数
#                 print('已等待消失:<%s>,大约等待(%.2f)s'%(objectName,round((NewTime-ActionTime),2)))
                break
            elif NewTime >= ActionTime+WaitTime:
                print('等待超时:<%s>,大约等待(%.2f)s'%(objectName,WaitTime))
                break
            else:
                time.sleep(0.5)
        
    def objTextCheck(self,objectName,DataName):  
        #检查文本内容是否符合预期
        CheckObject = self.getObject(objectName)
        ObjectText = CheckObject.get_attribute('text')
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
        print('验证成功，结果为<%s>'%ObjectText)

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
    
    
###_--------------------------滑动操作-----------------------------#

    def getWindowSize(self):
        x = self.get_window_size()['width']
        y = self.get_window_size()['height']
        return(x,y)
    
    def App_swipeUp(self,swipePoint = 0.5,SwipeStart = 0.8,swipeLenth = 0.5,swipeTime = 500):
        #swipeTime滑动间隔时间，单位毫秒(滑动速度)
        #swipeLenth滑动的跨幅(默认50%)，小数表示。
        #由于App边框会存在宽度或者高度
        #所以移动控制为半屏幕(0.75-0.25)
        #上滑则X不变，Y向上移动半个屏高。
        x,y = self.getWindowSize()
        x1=int(x*float(swipePoint))
        SwipeStart = 0.8
        y1=int(y*SwipeStart)
        y2=int(y*(SwipeStart-swipeLenth))
        self.swipe(x1,y1,x1,y2,swipeTime)
        
    def App_swipeDown(self,swipePoint = 0.5,swipeTime = 500):
        #swipeTime滑动间隔时间，单位毫秒
        #由于App边框会存在宽度或者高度
        #所以移动控制为半屏幕(0.75-0.25)
        #下滑则X不变，Y向下移动半个屏高。
        x,y = self.getWindowSize()
        x1=int(x*float(swipePoint))
        y1=int(y*0.25)
        y2=int(y*0.75)
        self.swipe(x1,y1,x1,y2,swipeTime)
        
    def App_swipeLeft(self,swipeTime = 500):
        #swipeTime滑动间隔时间，单位毫秒
        #由于App边框会存在宽度或者高度
        #所以移动控制为半屏幕(0.75-0.25)
        #左滑则y不变，x向左移动半个屏高
        x,y = self.getWindowSize()
        x1=int(x*0.75)
        y1=int(y*0.5)
        x2=int(y*0.25)
        self.swipe(x1,y1,x2,y1,swipeTime)
        
    def App_swipeRight(self,swipeTime):
        #swipeTime滑动间隔时间，单位毫秒
        #由于App边框会存在宽度或者高度
        #所以移动控制为半屏幕(0.75-0.25)
        #右滑则y不变，x向右移动半个屏高
        x,y = self.getWindowSize()
        x1=int(x*0.25)
        y1=int(y*0.5)
        x2=int(y*0.75)
        self.swipe(x1,y1,x2,y1,swipeTime)
        
    def App_Scroll(self,objA,objB):
        #操作模拟滑动动作
        #从objA的位置滑至objB
        #scroll方法需传入对象为2个定位的对象结果
        self.scroll(self.getObject(objA), self.getObject(objB))
        sleep(0.5)
        
    def ObjectLocationWithSwipe(self,objectName,Point = 0.5,Lenth = 0.5,SwipeType = 'up',WaitTime = 30):
        #Point默认滑动点，输入值代表屏幕宽度百分比(是点击左侧向上滑动，还是中间向上滑动)
        #SwipeType 滑动类型(up,down,left,right.上下左右)
        #WaitTime  滑动寻找对象的最大时间(默认30S,单次滑动耗时默认默认为500ms)
        ActionTime = time.time()
        WaitTime = int(WaitTime)
        IsExist = False
        Element = None
        while True:
            try:
                Element = self.getObject(objectName)
                #统一返回一个结果
                IsExist = True
            except:
                if SwipeType.lower() == 'up':
                    self.App_swipeUp(swipePoint = Point,
                                     swipeLenth = Lenth)
                elif SwipeType.lower() == 'down':
                    self.App_swipeDown(swipePoint = Point)
                elif SwipeType.lower() == 'left':
                    self.App_swipeLeft(swipePoint = Point)
                elif SwipeType.lower() == 'right':
                    self.App_swipeRight(swipePoint = Point)
                sleep(0.3)
            NewTime = time.time()
            if IsExist == True:
                #时间保留2位小数
#                 print('已完成滑动定位:<%s>,等待(%.2f)s'%(objectName,round((NewTime-ActionTime),2)))
                break
            elif NewTime >= ActionTime+WaitTime:
                print('滑动点击定位:<%s>,等待(%.2f)s'%(objectName,WaitTime))
                break
            else:
                time.sleep(0.5)
        return Element
    
    def ObjectsLocationWithSwipe(self,objectName,Point = 0.5,Lenth = 0.5,SwipeType = 'up',WaitTime = 30):
        #Point默认滑动点，输入值代表屏幕宽度百分比(是点击左侧向上滑动，还是中间向上滑动)
        #SwipeType 滑动类型(up,down,left,right.上下左右)
        #WaitTime  滑动寻找对象的最大时间(默认30S,单次滑动耗时默认默认为500ms)
        ActionTime = time.time()
        WaitTime = int(WaitTime)
        IsExist = False
        Element = None
        while True:
            try:
                Element = self.getObjects(objectName)
                #统一返回一个结果
                IsExist = True
            except:
                if SwipeType.lower() == 'up':
                    self.App_swipeUp(swipePoint = Point,
                                     swipeLenth = Lenth)
                elif SwipeType.lower() == 'down':
                    self.App_swipeDown(swipePoint = Point)
                elif SwipeType.lower() == 'left':
                    self.App_swipeLeft(swipePoint = Point)
                elif SwipeType.lower() == 'right':
                    self.App_swipeRight(swipePoint = Point)
                sleep(0.3)
            NewTime = time.time()
            if IsExist == True:
                #时间保留2位小数
#                 print('已完成滑动定位:<%s>,等待(%.2f)s'%(objectName,round((NewTime-ActionTime),2)))
                break
            elif NewTime >= ActionTime+WaitTime:
                print('滑动点击定位:<%s>,等待(%.2f)s'%(objectName,WaitTime))
                break
            else:
                time.sleep(0.5)
        return Element
    
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