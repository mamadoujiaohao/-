"""匯入模組"""

import pandas as pd
import datetime
import numpy as np

np.set_printoptions(threshold=np.inf)  #讓np可以print出全部的資訊
import pathlib

"""連接NAS"""
#from synology_dsm import SynologyDSM
#api = SynologyDSM("xxx", "xxx", "xxx", "xxx") #連接NAS(用主管的ID)

try:  #保護機制
  """Intro"""
  '''匯入檔案路徑'''
  print("---茂源匯單(網路訂購)---")

  #匯入明細路徑
  path明細 = pathlib.PureWindowsPath(
    input("拖曳明細檔案至此:"))  #不能用反斜線(NAS路徑是反斜線),所以使用pathlib
  path明細 = path明細.as_posix()  #用pathlib轉換反斜線為斜線
  path明細 = path明細.replace('"', '')

  #匯入工作表四路徑
  path表四 = pathlib.PureWindowsPath(
    input("拖曳工作表四至此:"))  #不能用反斜線(NAS路徑是反斜線),所以使用pathlib
  path表四 = path表四.as_posix()  #用pathlib轉換反斜線為斜線
  path表四 = path表四.replace('"', '')

  #讀入明細檔
  if path明細[0] == "/":
    OriginalFile = pd.read_excel("/%s" % (path明細))
  else:
    OriginalFile = pd.read_excel("%s" % (path明細))

  #讀入工作表四
  if path表四[0] == "/":
    Sheet4 = pd.read_excel("/%s" % (path表四))
  else:
    Sheet4 = pd.read_excel("%s" % (path表四))

  #準備欲儲存檔案之位置
  if path明細[0] == "/":
    writer = pd.ExcelWriter("/%s" % (path明細))
  else:
    writer = pd.ExcelWriter("%s" % (path明細))

  #保留明細
  明細 = pd.DataFrame(OriginalFile)
  明細 = 明細.sort_values(by=['出貨單號', '訂單序號'])  #將明細依照"出貨單號"、"訂單序號"做排序
  明細 = 明細.set_index("出貨單號")
  明細.to_excel(writer, sheet_name="大樹訂購明細")
  '''修主檔備註'''

  sheet2A = OriginalFile["出貨單號"]
  sheet2B = OriginalFile["備註"]

  sheet2 = {"大樹單號": sheet2B, "大樹備註": sheet2A}
  df_1 = pd.DataFrame(sheet2)  #建立表格sheet2
  df_1 = df_1.drop_duplicates(ignore_index=True)  #移除重複
  #df_1 = df_1['大樹備註'].fillna(value = "茂源未備註")
  df_1 = df_1[df_1['大樹單號'].str.contains("P")]  #移除"備註"無P之row
  df_1['大樹備註'] = "請購需求產生-" + df_1["大樹備註"]  #備註開頭加上"請購需求產生-"
  df_1 = df_1.set_index("大樹單號")
  df_1.to_excel(writer, sheet_name='修主檔備註')  #建立新工作表"修主檔備註",並匯入需求欄位
  '''建立工作表1'''
  #建立工作表1
  Sheet1 = pd.DataFrame(OriginalFile)

  #移除出貨量為0之列
  for i in range(0, len(Sheet1['出貨數量'])):
    if Sheet1['出貨數量'][i] == 0:
      Sheet1 = Sheet1.drop(index=i)

  #只留網路訂購(放這裡才不會打亂index))
  Sheet1 = Sheet1[Sheet1['備註'].str.contains("網路訂購")]

  #結合"網路訂購"與出貨單號
  Sheet1['備註'] = "網路訂購"
  Sheet1['出貨單號'] = Sheet1['備註'] + Sheet1['出貨單號']
  Sheet1_Temporary = {"備註": Sheet1["出貨單號"], "庫別": Sheet1["店號"]}
  Sheet1_Temporary_Array = pd.DataFrame(Sheet1_Temporary)

  #移除重複出貨單號
  Sheet1_Temporary_Array = Sheet1_Temporary_Array.drop_duplicates()

  #輸入相關資訊
  po_no = str(input("請輸入該採購日最終採購單號:"))
  updid = str(input("請輸入員編(五碼):"))
  po_date = str(input("請輸入採購日期(貼入樺穎日期):"))

  #單號等差遞增
  po_no = po_no[1:]
  po_no_int = int(po_no)
  po_no_list = []
  for i in range(1, len(Sheet1_Temporary_Array) + 1):
    Number = po_no_int + i
    po_no_list.append("P" + str(Number).zfill(len(str(po_no))))

  #時間格式要正確
  DateTemporary表1 = po_date[:10]
  正確時間 = datetime.datetime.strptime(DateTemporary表1, "%Y/%m/%d")

  #建立工作表1
  FirstSheetHeader = {
    "單號": po_no_list,
    "備註": Sheet1_Temporary_Array["備註"],
    "庫別": Sheet1_Temporary_Array["庫別"],
    "廠編": "Z05136",
    "日期": 正確時間,
    "採購": updid
  }
  FirstSheet = pd.DataFrame(FirstSheetHeader)
  FirstSheet = FirstSheet.set_index("單號")

  #庫別掐頭去尾
  for i in range(len(FirstSheet['庫別'])):
    FirstSheet['庫別'][i] = FirstSheet['庫別'][i][1:5]

  FirstSheet.to_excel(writer, sheet_name="3")  #出表1
  '''建立工作表2'''

  #建立單號
  Sheet2採購項目數 = Sheet1.reset_index()
  po_no_int2 = int(po_no)
  Sheet2單號 = []
  for i in range(len(Sheet2採購項目數)):
    if i > 0 and Sheet2採購項目數['出貨單號'][i - 1] == Sheet2採購項目數['出貨單號'][i]:
      Sheet2單號.append("P" + str(po_no_int2).zfill(len(str(po_no))))
    else:
      po_no_int2 = po_no_int2 + 1
      Sheet2單號.append("P" + str(po_no_int2).zfill(len(str(po_no))))

  Sheet2_Temporary = {
    "單號": Sheet2單號,
    "茂源碼": Sheet2採購項目數["茂源碼"]
  }  #, "商品代號":, "採購單價":, "採購數量":, "到貨期限":,}
  SecondSheet_Temporary = pd.DataFrame(Sheet2_Temporary)

  #建立品項序號
  Sheet2品項_單號_序號 = {
    "Sheet2品項_單號": Sheet2採購項目數["出貨單號"],
    "Sheet2品項_單號": Sheet2採購項目數["訂單序號"]
  }
  x = 1
  Sheet2序號 = []
  for i in range(len(Sheet2採購項目數["出貨單號"])):
    if i == 0:
      Sheet2序號.append(1)
    elif i > 0 and Sheet2採購項目數["出貨單號"][i] == Sheet2採購項目數["出貨單號"][i - 1]:
      x += 1
      Sheet2序號.append(x)
    elif i > 0 and Sheet2採購項目數["出貨單號"][i] != Sheet2採購項目數["出貨單號"][i - 1]:
      x = 1
      Sheet2序號.append(x)

  #商品代號轉換(茂源碼轉大數碼)
  Sheet4_Array = pd.DataFrame(Sheet4)
  Sheet2代號轉換 = pd.merge(SecondSheet_Temporary,
                        Sheet4_Array,
                        how='left',
                        left_on='茂源碼',
                        right_on='茂源碼')
  Sheet2代號轉換['大樹碼'][i] = '%.6d' % Sheet2代號轉換['大樹碼'][i]

  #大樹碼須為六碼,前面加000000
  Sheet2代號轉換['大樹碼'] = Sheet2代號轉換['大樹碼'].fillna(value=0)
  for i in range(len(Sheet2代號轉換['大樹碼'])):
    Sheet2代號轉換['大樹碼'][i] = int(Sheet2代號轉換['大樹碼'][i])
    Sheet2代號轉換['大樹碼'][i] = '%.6d' % Sheet2代號轉換['大樹碼'][i]

  #設定到貨期限
  DateTemporary = po_date[:10]
  Due_date = datetime.datetime.strptime(
    DateTemporary, "%Y/%m/%d") + datetime.timedelta(days=7)

  #設定採購單價
  Sheet2採購單價 = []
  for i in range(len(Sheet2採購項目數)):
    Sheet2採購單價.append(Sheet2採購項目數["出貨金額"][i] /
                      (Sheet2採購項目數["出貨數量"][i] * Sheet2代號轉換["1茂源=X大樹"][i]))

  #設定採購數量
  Sheet2採購數量 = []
  for i in range(len(Sheet2採購項目數)):
    Sheet2採購數量.append(Sheet2採購項目數["出貨數量"][i] * Sheet2代號轉換["1茂源=X大樹"][i])

  #建立工作表2
  SecondSheetHeader = {
    "單號": Sheet2代號轉換["單號"],
    "序號": Sheet2序號,
    "商品代號": Sheet2代號轉換["大樹碼"],
    "採購單價": Sheet2採購單價,
    "採購數量": Sheet2採購數量,
    "到貨期限": Due_date
  }

  if not len(Sheet2代號轉換["單號"]) == len(Sheet2序號) == len(
      Sheet2代號轉換["大樹碼"]) == len(Sheet2採購單價) == len(Sheet2採購數量):
    print(
      "--------------------------------------------------------------------------"
    )
    print("表四茂源碼有重複,請移除重複項")
    print(len(Sheet2代號轉換["單號"]), len(Sheet2序號), len(Sheet2代號轉換["大樹碼"]),
          len(Sheet2採購單價), len(Sheet2採購數量))

  SecondSheet = pd.DataFrame(SecondSheetHeader)
  SecondSheet = SecondSheet.set_index("單號")
  SecondSheet.to_excel(writer, sheet_name="4")

except:
  try:
    writer.close()
  except:
    print()
  print("\n\n")
  print(
    "--------------------------------------------------------------------------"
  )
  input("匯入的資料有誤,請先檢查明細和剛剛輸入之資料(單號、員編、日期)\n請按Enter鍵結束")
'''關閉writer'''
writer.close()

#查看表四有無該品項
茂源碼_表四缺 = []
茂源品名_表四缺 = []
茂源單位_表四缺 = []
for i in range(len(Sheet1["茂源碼"])):
  if SecondSheet['商品代號'][i] == "000000":
    茂源碼_表四缺.append(Sheet1["茂源碼"][i:i + 1].item())
    茂源品名_表四缺.append(Sheet1["品名"][i:i + 1].item())
    茂源單位_表四缺.append(Sheet1["單位"][i:i + 1].item())

表四缺少品項 = {"茂源碼": 茂源碼_表四缺, "茂源單位": 茂源單位_表四缺, "茂源品名": 茂源品名_表四缺}
表四缺少品項_DataFrame = pd.DataFrame(表四缺少品項)

print("-------------------------------------")
print("品項對照表缺少品項:")
if 表四缺少品項_DataFrame.empty == True:
  print("無")
  input("\n請按Enter鍵結束\n\n")
else:
  print(表四缺少品項_DataFrame)
  print("-------------------------------------")
  print("\n上方為對照表缺少品項\n請於對照表新增該品項後重新跑一次")
  input("\n請按Enter鍵結束\n\n")
