from datetime import date
import openpyxl
from ping3 import ping
import requests
from requests import Timeout

# 調査対象（Excelファイル）を指定
wb_source=openpyxl.load_workbook('destinations/dest.xlsx')
ws_source=wb_source['Sheet1']

# レポート（Excelファイル）のファイル名を指定
filename_report='reports/report_' + date.today().strftime('%Y%m%d') + '.xlsx'
openpyxl.Workbook().save(filename_report)

# データを2次元配列に格納
PingList=[[0 for i in range(ws_source.max_column)] for j in range(ws_source.max_row)]
for x in range(0, ws_source.max_row):
    for y in range(0, ws_source.max_column):
        PingList[x][y]=ws_source.cell(row=x+1, column=y+1).value

# レポート（Excelファイル）を読み込み
wb_report=openpyxl.load_workbook(filename_report)
ws_report=wb_report['Sheet']

# 調査対象の数だけループ処理を開始
for i in range(0, ws_source.max_row):        
    # 調査対象のIPアドレスを書き込み
    ws_report.cell(row=i+2, column=1).value=PingList[i][0]

    # ping応答時間をセルに書き込み（pingをj回送信する）
    for j in range(4):
        #★ping送信
        ping_result=ping(PingList[i][0], unit='ms')
        #ping送信回数だけセルに書き込み
        ws_report.cell(row=1, column=j+2).value=str(j+1) + '回目'
        ws_report.cell(row=i+2, column=j+2).value=ping_result
        #print(ping_result)

    #★HTTPリクエスト送信
    try:
        http_result=requests.get('http://' + PingList[i][0], timeout=3.0).status_code
    except Timeout:
        http_result='TimeOut'
        pass
    #HTTPレスポンスをセルに書き込み
    ws_report.cell(row=i+2, column=j+3).value=http_result

# ヘッダー情報をセルに書き込み
ws_report.cell(row=1, column=1).value = 'IPアドレス'
ws_report.cell(row=1, column=j+3).value = 'HTTPレスポンス'

# 列幅を変更
ws_report.column_dimensions['A'].width=15

# レポート（Excelファイル）に保存
wb_report.save(filename_report)