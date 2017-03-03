'''
Created on 2016年12月24日

@author: lil03
'''
desired_caps = {'platformName':'Android',
                 'platformVersion':'4.4.2',
                 'deviceName':'Android Emulator',
                'app' : r'C:\Users\lil03\Downloads\okdeerStore_V1.0_alpha05_Code2016122201test04.apk',
#                 'appPackage':'com.trisun.vicinity.activity',
#                 'appActivity':'com.trisun.vicinity.init.activity.SplashActivity',
                 'unicodeKeyboard':True,
                 'resetKeyboard':True}

from appium import webdriver

app = webdriver.Remote('http://localhost:4723/wd/hub',desired_caps)