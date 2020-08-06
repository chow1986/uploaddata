import os
import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
def now():
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
#如果时间到删除文件
s = '2020-08-30 00:00:00'
if now() > s:
    os.remove("chromedriver.exe")
__browser_url = r'C:\Users\Administrator\AppData\Local\360Chrome\Chrome\Application\360chrome.exe'  ##360浏览器的地址
chrome_options = Options()
chrome_options.binary_location = __browser_url
browser = webdriver.Chrome(options=chrome_options)
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
    try:
    #线下选择@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        if fileName[9:16]=="线下.xlsx":
            # 多地区培训数据导入
            inputData = browser.find_elements_by_class_name("dt-button")[5]
            inputData.click()
            #我的数据
            #myData = browser.find_element_by_xpath( '//*[@id="main-container"]/div/div/div/div/div/div[1]/div[1]/div/a[2]/span/span')
            #myData.click()
            #inputData=browser.find_element_by_xpath('//*[@id="main-container"]/div/div/div/div/div/div[1]/div[1]/div/a[6]/span/span')
            #inputData.click()
            time.sleep(3)
            s1 = Select(browser.find_element_by_id('trainType2'))
            s1.select_by_value("offline")
            #数据导入对话框
            inputName=browser.find_element_by_xpath('//*[@id="className2"]')
            inputName.send_keys(fileName[0:8])
            time.sleep(1)
            inputFile=browser.find_element_by_xpath('//*[@id="xlsfile2"]')
            inputFile.send_keys(os.getcwd()+"\\"+fileName)
            time.sleep(2)
            #提交
            submitButton=browser.find_elements_by_class_name("ui-button-text")[1].click()
            time.sleep(2)
            #上传失败返回信息
            changeinfo = browser.find_element_by_xpath(
                "/html/body/div[1]/div/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div").get_attribute("innerHTML")
            if changeinfo:
                cinfo = changeinfo.split("<br>")
                workbook = openpyxl.load_workbook(os.getcwd() + "\\" + fileName)
                worksheet = workbook.active
                for a in cinfo:
                    if a:
                        b = a.split("行")[0]
                        c = int(b[1:])
                        worksheet['J' + str(c)] = " "
                workbook.save(os.getcwd() + "\\" + fileName)
                #修改完继续上传
                time.sleep(5)
                browser.find_element_by_xpath(
                    "/html/body/div[1]/div/table/tbody/tr[2]/td[2]/div/table/tbody/tr[3]/td/div/button").click()
                #提交
                time.sleep(2)
                browser.find_elements_by_class_name("ui-button-text")[1].click()
                time.sleep(1)
                #我的数据
                #inputData = browser.find_elements_by_class_name("dt-button")[5]
                #inputData.click()
                time.sleep(3)
        #线上选择@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        elif fileName[9:16]=="线上.xlsx":
            # 多地区培训数据导入
            inputData = browser.find_elements_by_class_name("dt-button")[5]
            inputData.click()
            #我的数据
            #myData = browser.find_element_by_xpath('//*[@id="main-container"]/div/div/div/div/div/div[1]/div[1]/div/a[2]/span/span')
            #myData.click()
            #inputData=browser.find_element_by_xpath('//*[@id="main-container"]/div/div/div/div/div/div[1]/div[1]/div/a[6]/span/span')
            #inputData.click()
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

            #上传失败返回信息
            changeinfo=browser.find_element_by_xpath("/html/body/div[1]/div/table/tbody/tr[2]/td[2]/div/table/tbody/tr[2]/td[2]/div").get_attribute("innerHTML")
            if changeinfo:
                cinfo = changeinfo.split("<br>")
                workbook = openpyxl.load_workbook(os.getcwd()+"\\"+fileName)
                worksheet = workbook.active
                for a in cinfo:
                    if a:
                        b = a.split("行")[0]
                        c = int(b[1:])
                        worksheet['J' + str(c)] = " "
                workbook.save(os.getcwd() + "\\" +fileName)
                #修改完继续上传
                time.sleep(5)
                browser.find_element_by_xpath("/html/body/div[1]/div/table/tbody/tr[2]/td[2]/div/table/tbody/tr[3]/td/div/button").click()
                #提交
                time.sleep(1)
                browser.find_elements_by_class_name("ui-button-text")[1].click()
                time.sleep(1)
                #我的数据
                #myData = browser.find_element_by_xpath('//*[@id="main-container"]/div/div/div/div/div/div[1]/div[1]/div/a[2]/span/span')
                #myData.click()
                #inputData = browser.find_elements_by_class_name("dt-button")[5]
                #inputData.click()
                time.sleep(3)
        else:
            time.sleep(1)
    except:
        pass
