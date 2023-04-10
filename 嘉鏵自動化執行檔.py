"""匯入模組"""

import pandas as pd
import datetime
import numpy as np

np.set_printoptions(threshold=np.inf)  #讓np可以print出全部的資訊
import pathlib
"""連接NAS"""
#from synology_dsm import SynologyDSM
#api = SynologyDSM("xxx", "xxx", "xxx", "xxx")#連接NAS(用主管的ID)
"""爬蟲(暫時無)"""

try:  #保護檔案
  """Intro"""
  '''匯入檔案路徑'''
  print("------嘉鏵匯單------")

  #匯入明細路徑
  path明細 = pathlib.PureWindowsPath(
    input("拖曳明細檔案至此:"))  #不能用反斜線(NAS路徑是反斜線),所以使用pathlib
  path明細 = path明細.as_posix()  #用pathlib轉換反斜線為斜線
  path明細 = path明細.replace('"', '')

  #匯入對照表路徑
  path對照表 = pathlib.PureWindowsPath(
    input('拖曳"嘉鏵品項與店編對照表"至此:'))  #不能用反斜線(NAS路徑是反斜線),所以使用pathlib
  path對照表 = path對照表.as_posix()  #用pathlib轉換反斜線為斜線
  path對照表 = path對照表.replace('"', '')

  #輸入相關資訊
  單號 = str(input("請輸入該採購日最終採購單號:"))
  採購人員 = str(input("請輸入員編(五碼):"))
  採購日期 = str(input("請輸入採購日期(貼入樺穎日期):"))

  #讀入明細檔
  if path明細[0] == "/":
    明細 = pd.ExcelFile("/%s" % (path明細))
  else:
    明細 = pd.ExcelFile("%s" % (path明細))

  第一頁明細 = pd.DataFrame(明細.parse(0))
  第二頁明細 = pd.DataFrame(明細.parse(1))

  #讀入對照表
  if path對照表[0] == "/":
    對照表 = pd.ExcelFile("/%s" % (path對照表))
  else:
    對照表 = pd.ExcelFile("%s" % (path對照表))

  店編對照表 = 對照表.parse(0)
  品項對照表 = 對照表.parse(1)

  #準備欲儲存檔案之位置
  if path明細[0] == "/":
    主檔 = pd.ExcelWriter("/%s" % (path明細))
  else:
    主檔 = pd.ExcelWriter("%s" % (path明細))

  #保留明細(Sheet0、Sheet1)
  第一頁明細_ResetIndex = 第一頁明細.set_index("發票號碼")
  第二頁明細_ResetIndex = 第二頁明細.set_index("發票號碼")
  第一頁明細_ResetIndex.to_excel(主檔, sheet_name="發票明細1")
  第二頁明細_ResetIndex.to_excel(主檔, sheet_name="發票明細2")
  """製作FirstSheet"""
  #處理單號
  FirstSheet單號 = 單號[1:]
  FirstSheet單號_數字 = int(FirstSheet單號)
  FirstSheet單號list = []
  for i in range(1, len(第一頁明細) + 1):
    Number = FirstSheet單號_數字 + i
    FirstSheet單號list.append("P" + str(Number).zfill(len(str(單號)) - 1))

  #處理備註
  備註list = []
  for i in range(len(第一頁明細)):
    備註list.append("發票號碼" + 第一頁明細["發票號碼"][i])

  #處理庫別
  買方名稱 = 第一頁明細["買方名稱"]
  買方名稱轉庫別 = pd.merge(買方名稱, 店編對照表, how="left", left_on="買方名稱", right_on="買方名稱")

  買方名稱轉庫別["店代號"] = 買方名稱轉庫別["店代號"].fillna(value=0)  #使庫別維持四碼
  for i in range(len(買方名稱轉庫別["店代號"])):
    買方名稱轉庫別["店代號"][i] = '%.4d' % int(買方名稱轉庫別["店代號"][i])

  #處理廠編(不用)

  #處理日期
  採購日期_日期 = 採購日期[:10]
  正確時間 = datetime.datetime.strptime(採購日期_日期, "%Y/%m/%d")

  #處理採購(不用)

  #組合FirstSheet
  FirstSheet = {
    "單號": FirstSheet單號list,
    "備註": 備註list,
    "庫別": 買方名稱轉庫別["店代號"],
    "廠編": "Z02766",
    "日期": 正確時間,
    "採購": 採購人員
  }
  FirstSheetDataFrame = pd.DataFrame(FirstSheet)
  FirstSheetDataFrame = FirstSheetDataFrame.set_index("單號")
  FirstSheetDataFrame.to_excel(主檔, sheet_name="3")
  """製作SecondSheet"""
  #處理單號
  SecondSheet單號_數字 = int(單號[1:])
  SecondSheet單號 = []

  for i in range(len(第二頁明細)):
    if i > 0 and 第二頁明細["發票號碼"][i] == 第二頁明細["發票號碼"][i - 1]:
      SecondSheet單號.append("P" + str(SecondSheet單號_數字).zfill(len(str(單號)) - 1))
    else:
      SecondSheet單號_數字 = SecondSheet單號_數字 + 1
      SecondSheet單號.append("P" + str(SecondSheet單號_數字).zfill(len(str(單號)) - 1))

  #處理序號(不用)

  #處理商品代號
  #字尾不可去除項
  字尾不可去除項 = ["BETAMETHASONE 貝他每松軟膏【花",
             "BETAMETHASONE 貝他每松軟膏【清"]  #這個list可自行新增項目
  品名去除最後字元 = 第二頁明細
  for i in range(len(品名去除最後字元["品名"])):
    if str(品名去除最後字元["品名"][i]) in 字尾不可去除項:
      品名去除最後字元["品名"][i] = 品名去除最後字元["品名"][i]
    else:
      品名去除最後字元["品名"][i] = 品名去除最後字元["品名"][i][:-1]

  SecondSheet品項對接 = pd.merge(品名去除最後字元,
                             品項對照表,
                             how="left",
                             left_on=['品名', '單位'],
                             right_on=['品名(去除最後一字元)', '單位'])
  SecondSheet商品代號 = SecondSheet品項對接["大樹碼"]
  SecondSheet商品代號 = SecondSheet商品代號.fillna(value=0)
  for i in range(len(SecondSheet商品代號)):
    SecondSheet商品代號[i] = int(SecondSheet商品代號[i])
    SecondSheet商品代號[i] = '%.6d' % SecondSheet商品代號[i]

  #處理到貨期限
  採購日期_日期 = 採購日期[:10]
  SecondSheet到貨期限 = datetime.datetime.strptime(
    採購日期_日期, "%Y/%m/%d") + datetime.timedelta(days=7)

  #處理採購數量(承接#處理商品代號)
  SecondSheet採購數量 = []
  for i in range(len(SecondSheet品項對接["數量"])):
    SecondSheet採購數量.append(str(SecondSheet品項對接["數量"][i]))

  SecondSheet採購數量 = [i.replace(",", "")
                     for i in SecondSheet採購數量]  #不知道為什麼不能跟上面的for迴圈合併
  SecondSheet採購數量array = np.array(SecondSheet採購數量)
  SecondSheet採購數量array = SecondSheet採購數量array.astype(np.float)

  for i in range(len(SecondSheet採購數量)):
    SecondSheet採購數量array[
      i] = SecondSheet採購數量array[i] * SecondSheet品項對接["入數"][i]

  #處理採購單價(承接##處理採購數量)
  SecondSheet嘉鏵單價_字串 = []
  for i in range(len(SecondSheet品項對接["單價"])):
    SecondSheet嘉鏵單價_字串.append(SecondSheet品項對接["單價"][i])

  SecondSheet嘉鏵單價_字串 = [i.replace(",", "") for i in SecondSheet嘉鏵單價_字串]
  SecondSheet採購單價 = np.array(SecondSheet嘉鏵單價_字串)
  SecondSheet採購單價 = SecondSheet採購單價.astype(np.float)

  for i in range(len(SecondSheet採購單價)):
    SecondSheet採購單價[i] = (SecondSheet採購單價[i] / SecondSheet品項對接["入數"][i]) * 1.05

  #組合SecondSheet
  SecondSheet = {
    "單號": SecondSheet單號,
    "序號": 第二頁明細["序號"],
    "商品代號": SecondSheet商品代號,
    "採購單價": SecondSheet採購單價,
    "採購數量": SecondSheet採購數量array,
    "到貨期限": SecondSheet到貨期限
  }
  SecondSheetDataFrame = pd.DataFrame(SecondSheet)
  SecondSheetDataFrame = SecondSheetDataFrame.set_index("單號")
  SecondSheetDataFrame.to_excel(主檔, sheet_name="4")

except:  #保護檔案
  try:  #如果根本沒匯入file,直接關file會有bug   #如果遇到bug,後面的file可能會沒關到,所以要一個一個分開關閉
    明細.close()
  except:
    print()
  try:
    對照表.close()
  except:
    print()
  try:
    主檔.close()
  except:
    print()
  print(
    "--------------------------------------------------------------------------"
  )
  input("匯入的資料有誤,請先檢查剛剛匯入之檔案和輸入之資料(單號、員編、日期)\n請按Enter鍵結束")
'''關閉writer'''

主檔.close()
明細.close()
對照表.close()
"""確認表三店編皆有"""

缺少店編 = []
for i in range(len(FirstSheet["庫別"])):
  if FirstSheet["庫別"][i] == "0000":
    缺少店編.append(買方名稱轉庫別["買方名稱"][i])

缺少店編_dic = {"名稱": 缺少店編}
缺少店編_DataFrame = pd.DataFrame(缺少店編_dic)

print("\n\n")
print(
  "--------------------------------------------------------------------------")
print('!!!若有"大樹醫藥股份有限公司(桃園)+印章",請至明細檔內刪除該筆交易明細!!!')
print("請新增以下店編:")

if 缺少店編_DataFrame.empty == True:
  print("無")
else:
  print(缺少店編_DataFrame)

print(
  "--------------------------------------------------------------------------")
"""確認表四品項皆有"""

嘉鏵品名_表四缺 = []
嘉鏵單位_表四缺 = []
for i in range(len(SecondSheet商品代號)):
  if SecondSheet商品代號[i] == "000000":
    嘉鏵品名_表四缺.append(品名去除最後字元["品名"][i:i + 1].item())
    嘉鏵單位_表四缺.append(品名去除最後字元["單位"][i:i + 1].item())

表四缺少品項 = {"嘉鏵單位": 嘉鏵單位_表四缺, "嘉鏵品名(字尾-1)": 嘉鏵品名_表四缺}
表四缺少品項_DataFrame = pd.DataFrame(表四缺少品項)

print("\n")
print(
  "--------------------------------------------------------------------------")
print("請新增以下品項:")
if 表四缺少品項_DataFrame.empty == True:
  print("無")
else:
  print(表四缺少品項_DataFrame)

print(
  "--------------------------------------------------------------------------")
"""outro"""

if 表四缺少品項_DataFrame.empty == True and 缺少店編_DataFrame.empty == True:
  input("\n請按Enter鍵繼續\n\n")
else:
  print("\n上方為對照表缺少店編&品項\n請於對照表新增後重新跑一次")
  input("\n請按Enter鍵繼續\n\n")
"""無法正確merge(葉酸錠)"""
葉酸錠列數 = []

if len(品名去除最後字元[品名去除最後字元["品名"] == "FOLIC ACID 5MG 葉酸錠 10T/排"]) > 0:  #品名結尾第一種
  葉酸錠列數.append(
    str(品名去除最後字元[品名去除最後字元["品名"] == "FOLIC ACID 5MG 葉酸錠 10T/排"].index.values +
        2))

if len(品名去除最後字元[品名去除最後字元["品名"] == "FOLIC ACID 5MG 葉酸錠 10T/排("]) > 0:  #品名結尾第二種
  葉酸錠列數.append(
    str(品名去除最後字元[品名去除最後字元["品名"] == "FOLIC ACID 5MG 葉酸錠 10T/排("].index.values +
        2))

if len(葉酸錠列數) > 0:
  print(
    "--------------------------------------------------------------------------"
  )
  [
    print(
      "第", i,
      "列為'葉酸錠',該品項對應多個大樹碼,由於嘉鏵提供的品名本程式無法判斷:\n請採購至該列確認大樹碼是否正確(120937榮民價:1.1&1.2 134537應元價:1.8&1.9)"
    ) for i in 葉酸錠列數
  ]
  print("(2022.12 單價)")
  print(
    "--------------------------------------------------------------------------"
  )
  input("\n請按Enter鍵結束\n\n")
