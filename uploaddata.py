import os
import time
from selenium import webdriver
from selenium.webdriver.support.ui import Select
def now():
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
#如果时间到删除文件
s = '2020-08-30 00:00:00'
if now() > s:
    os.remove("chromedriver.exe")
browser = webdriver.Chrome()
browser.get('http://px.zjsafety.gov.cn')
time.sleep(80)

#我的数据按钮
myData=browser.find_element_by_class_name("nav-list").find_elements_by_tag_name("li")[5]
myData.click()
time.sleep(3)
#inputData=browser.find_element_by_xpath('//*[@id="main-container"]/div/div/div/div/div/div[1]/div[1]/div/a[6]')
#切换框架
browser.switch_to.frame("iFrame4")
#文件选择
fileList = os.listdir(os.getcwd())
#遍历文件夹中所有文件
for fileName in fileList:
    #线下选择@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    if fileName[9:16]=="线下.xlsx":
        # 多地区培训数据导入
        inputData = browser.find_elements_by_class_name("glyphicon-import")[1]
        inputData.click()
        time.sleep(3)
        s1 = Select(browser.find_element_by_id('trainType2'))
        s1.select_by_value("offline")
        #数据导入对话框
        inputName=browser.find_element_by_xpath('//*[@id="className2"]')
        print(fileName[0:8])
        inputName.send_keys(fileName[0:8])
        time.sleep(1)

        inputFile=browser.find_element_by_xpath('//*[@id="xlsfile2"]')
        print(os.getcwd()+"\\"+fileName)
        inputFile.send_keys(os.getcwd()+"\\"+fileName)
        time.sleep(2)

        #提交
        submitButton=browser.find_elements_by_class_name("ui-button-text")[1].click()
        time.sleep(3)
#线上选择@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    elif fileName[9:16]=="线上.xlsx":
        # 多地区培训数据导入
        inputData = browser.find_elements_by_class_name("glyphicon-import")[1]
        inputData.click()
        time.sleep(3)
        s2 = Select(browser.find_element_by_id('trainType2'))
        s2.select_by_value("online")
        #数据导入对话框
        #inputName=browser.find_element_by_xpath('//*[@id="className2"]')
        #inputName.send_keys(fileName[0:8])
        #time.sleep(1)
        inputFile=browser.find_element_by_xpath('//*[@id="xlsfile2"]')
        inputFile.send_keys(os.getcwd()+"\\"+fileName)
        time.sleep(2)

        #提交
        submitButton=browser.find_elements_by_class_name("ui-button-text")[1].click()
        time.sleep(3)
    else:
        time.sleep(0.5)