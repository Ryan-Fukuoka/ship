import streamlit as st
import openpyxl
import pandas as pd
import datetime
import ship_normal
import ship_quick
import ship_package
import tutorial
from configparser import ConfigParser
from io import BytesIO

st.sidebar.image('logo.png', width=60)
st.sidebar.title('郵局包裹自動填寫系統')
st.sidebar.markdown("""---""")
orderFile = st.sidebar.file_uploader('🗂️ 上傳平台訂單CSV') # 上傳CSV 檔
today = datetime.datetime.now()
before = today - datetime.timedelta(days=7)
inputDate = st.sidebar.date_input('📅 訂單起迄日', (before, today))
service = st.sidebar.radio('✈️ 請選擇郵寄服務種類', ['國際平常小包', '國際包裹', '國際快捷郵件'])

try:
    strDate, endDate = inputDate
except ValueError:
    st.sidebar.error("請選擇結束日期")
    st.stop()

# 讀取產品列表products.csv
products = pd.read_csv('products.csv')

# 寄件人選擇清單
senderConf = ConfigParser()
senderConf.read('sender.ini')
sects = senderConf.sections()
sndSec = st.sidebar.selectbox('✉️ 請選擇要發送的寄件人',sects)

st.sidebar.markdown("""---""")
# 系統說明
tutorial.dropdown()

# 上傳平台訂單CSV
byteWb = BytesIO() # 將Excel 存在檔案串流中
if orderFile is not None:
	st.markdown("""---""")
	st.header('訂單內容')
	orderDf = pd.read_csv(orderFile)

	# 將年份格式YY 改為YYYY
	orderDf['Sale Date'] = orderDf['Sale Date'].str[0:6]+'20'+orderDf['Sale Date'].str[6:9]
	orderDf['Sale Date'] = pd.to_datetime(orderDf['Sale Date'], format='%m/%d/%Y')

	# 條件過濾後的訂單資料resOrderDf
	cond_1 = orderDf['Sale Date']>=strDate.strftime('%Y/%m/%d')
	cond_2 = orderDf['Sale Date'] <= endDate.strftime('%Y/%m/%d')
	cond_3 = orderDf['Date Shipped'].isnull()
	resOrderDf = orderDf[cond_1 & cond_2 & cond_3]
	resOrderDf
	st.markdown("""---""")

	# 將訂單和產品做join
	joinDf = pd.merge(resOrderDf, products, how='left', on='Item Name')
	orderIds = joinDf['Order ID'].unique()
	shipNames = joinDf['Ship Name'].unique()
	seltNames = st.multiselect('請選擇'+service+'收件人', shipNames, shipNames)

	# 輸入參數：訂單和產品DataFrame, 訂單中所有Order ID, 選取的收件人名字, 寄件人設定檔Section
	if service == '國際平常小包':
		wb = ship_normal.getExcel(joinDf, orderIds, seltNames, sndSec)
		wb.save(byteWb)
		st.download_button(label='下載郵局出貨單Excel', data=byteWb, 
			file_name='normal.xls', mime='application/vnd.ms-excel')
	elif service == '國際包裹':
		wb = ship_package.getExcel(joinDf, orderIds, seltNames, sndSec)
		wb.save(byteWb)
		st.download_button(label='下載郵局出貨單Excel', data=byteWb, 
			file_name='package.xls', mime='application/vnd.ms-excel')
	elif service == '國際快捷郵件':
		wb = ship_quick.getExcel(joinDf, orderIds, seltNames, sndSec)
		wb.save(byteWb)
		st.download_button(label='下載郵局出貨單Excel', data=byteWb, 
			file_name='quick.xls', mime='application/vnd.ms-excel')

else:
	st.markdown("""---""")
	st.subheader('🗂️ 請上傳CSV檔案')


