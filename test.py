import os
import openpyxl
changeinfo="第10行，未注册的乡镇街道：美国<br>第11行，未注册的乡镇街道：宁波大学<br>第12行，未注册的乡镇街道：美国<br>第16行，未注册的乡镇街道：美国<br>第23行，未注册的乡镇街道：美国<br>第29行，未注册的乡镇街道：宁波大学<br>"
changeinfo=changeinfo.split("<br>")
workbook = openpyxl.load_workbook(os.getcwd() + "\\" + "20200097-线上.xlsx")
worksheet = workbook.active
for a in changeinfo:
    '''
    b = a.split("行")[0]
    c = int(b[1:])
    print(c)
    worksheet['J' + c] = " "
    '''
    if a:
        b = a.split("行")[0]
        c = int(b[1:])
        print(c)
        worksheet['J' + str(c)] = "1234567 "
workbook.save(os.getcwd() + "\\" + "修改-20200097-线上.xlsx")





'''
changeinfo = changeinfo.split("行")[0]
changeinfo = int(changeinfo[1:])
# 打开excel文件修改数据
workbook = openpyxl.load_workbook(os.getcwd() + "\\" + "20200097-线上.xlsx")
worksheet = workbook.active
worksheet['J'+str(changeinfo)] = " "
workbook.save(os.getcwd() + "\\" + "修改-20200097-线上.xlsx")
'''