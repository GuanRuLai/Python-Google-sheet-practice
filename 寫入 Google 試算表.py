# -*- coding: utf-8 -*-
"""
Created on Thu Feb  9 21:43:07 2023

@author: HP
"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials as sac
# 設定金鑰檔包含檔名的路徑
auth_json="pythonconnectgsheet1-377312-57a7e15183e5.json"
# 設定程式可操作的範圍
gs_scopes=['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']
# 以金鑰檔及操作範圍設定憑證
cr=sac.from_json_keyfile_name(auth_json,gs_scopes)
# 取得操作憑證，建立連線物件
gc=gspread.authorize(cr)
# 以檔名開啟試算表
gsheet=gc.open("PythonConnectGSheet")
# 以 id 開啟試算表
gsheet=gc.open_by_key("1rA1zj82pEntaIpRkmYNKJ9USLjQ6p1Iud19MgYz8hKQ")
# 開啟工作簿
wks=gsheet.sheet1
# 清除所有內容
wks.clear()
# 新增列
listtitle=["座號","姓名","國文","英文","數學"]
wks.append_row(listtitle) # 標題
listdatas=[[1,"葉大雄",65,62,40],
           [2,"陳靜香",85,90,87],
           [3,"王聰明",92,90,95]]
for listdata in listdatas:
    wks.append_row(listdata) # 資料內容
    
