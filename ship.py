import streamlit as st
import openpyxl
import pandas as pd
import datetime
from configparser import ConfigParser
from io import BytesIO

st.sidebar.image('logo.png', width=60)
st.sidebar.title('郵局包裹自動填寫系統')
st.sidebar.markdown("""---""")
orderFile = st.sidebar.file_uploader('🗂️ 上傳平台訂單CSV') # 上傳CSV 檔
today = datetime.datetime.now()
before = today - datetime.timedelta(days=7)
inputDate = st.sidebar.date_input('📅 訂單起迄日', (before, today))
try:
    strDate, endDate = inputDate
except ValueError:
    st.sidebar.error("請選擇結束日期")
    st.stop()
st.sidebar.markdown("""---""")

# 讀取產品列表products.csv
products = pd.read_csv('products.csv')

# 讀取設定檔
config = ConfigParser()
config.read('setting.ini')

# 系統說明
option = st.sidebar.selectbox('系統說明',['請選擇','操作說明','setting.ini','products.csv'])
if option == '操作說明':
	st.subheader('操作說明')
	st.markdown("""---""")
	st.info('* 上傳商城平台訂單CSV檔 \n * 訂單起訖日：預設值時間區間是今天日期往前推7天')
	st.error('重要：因為郵局Excel 範例檔案包含程式，必須要在下載檔案後，\
		使用Excel 應用程式開啟，Excel 檔中的程式被執行之後，儲存Excel 檔案，最後才能將檔案上傳到郵局系統，\
		否則郵局系統會發生錯誤')
elif option == 'setting.ini':
	st.subheader('設定檔setting.ini')
	st.markdown("""---""")
	st.info('* 請在sender 輸入送件人資訊 \n * 因為商城與郵局國家命名規則不同，\
		請在country 設定兩份文件的國別，設定規則：商城文件中「Ship Country」欄位 = 郵局文件中「收件人國別」，\
		如果沒有設定，將顯示空白，例：United States=U.S.A.')
	st.markdown('[更新設定檔網址](https://github.com/Ryan-Fukuoka/ship/blob/main/setting.ini)')
elif option == 'products.csv':
	st.subheader('產品列表products.csv')
	st.info('* 請在CSV 檔中輸入產品資訊，欄位名稱依序是Item Name,length,width,high,weight,\
		content,currency,description,price\
		\n * 商城文件中的Item Name 必須要和products.csv 中的Item Name 相符，否則會無法對應到產品資訊\
		\n * 使用Excel 修改CSV 檔可能會造成中文編碼錯誤，請儘量使用純文字檔編輯器或其他軟體修改')
	st.markdown('[更新產品列表網址](https://github.com/Ryan-Fukuoka/ship/blob/main/products.csv)')
	products

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

	# 將訂單和產品做join，用Order ID 做Groupby
	joinDf = pd.merge(resOrderDf, products, how='left', on='Item Name')
	orderIds = joinDf['Order ID'].unique()

	
	
	# 將訂單資料寫入到郵局出貨單
	rowNum = 3 # 從第3列開始往下增加資料
	wb = openpyxl.load_workbook(filename='format.xlsx')
	ws = wb['託運清單']
	for orderId in orderIds:
		order = joinDf[joinDf['Order ID'] == orderId] # 使用Order ID 找到join 後的table 訂單
		ws['B'+str(rowNum)].value = order['Ship Name'].unique()[0]
		ws['F'+str(rowNum)].value = config['sender']['type']

		# 商城與郵局的國別命名不同，需要在setting.ini 設定對照表
		try:		
			ws['G'+str(rowNum)].value = config['country'][order['Ship Country'].unique()[0]]
		except KeyError:
			ws['G'+str(rowNum)].value = ''
		ws['H'+str(rowNum)].value = order['Ship Zipcode'].unique()[0]
		ws['I'+str(rowNum)].value = order['Ship State'].unique()[0]
		ws['J'+str(rowNum)].value = order['Ship City'].unique()[0]
		ws['K'+str(rowNum)].value = order['Ship Address1'].unique()[0]
		ws['P'+str(rowNum)].value = config['sender']['name']
		ws['R'+str(rowNum)].value = config['sender']['tel']
		ws['S'+str(rowNum)].value = config['sender']['zip']
		ws['T'+str(rowNum)].value = config['sender']['city']
		ws['U'+str(rowNum)].value = config['sender']['address']
		ws['V'+str(rowNum)].value = order['content'].unique()[0]
		ws['W'+str(rowNum)].value = order['weight'].sum()
		ws['X'+str(rowNum)].value = order['length'].sum()
		ws['Y'+str(rowNum)].value = order['width'].sum()
		ws['Z'+str(rowNum)].value = order['high'].sum()
		ws['AA'+str(rowNum)].value= order['currency'].unique()[0]
		ws['AB'+str(rowNum)].value= order['description'].unique()[0]
		ws['AC'+str(rowNum)].value= order['Quantity'].sum()
		ws['AD'+str(rowNum)].value= order['weight'].sum()
		ws['AE'+str(rowNum)].value= order['price'].unique()[0]
		ws['AF'+str(rowNum)].value= order['Quantity'].sum() * order['price'].unique()[0]
		rowNum +=1

	# 將處理好的Excel轉為二進位檔案串流下載
	wb.save(byteWb)
	st.download_button(label='下載郵局出貨單Excel', data=byteWb, 
		file_name='ship.xls', mime='application/vnd.ms-excel')
else:
	st.markdown("""---""")
	st.subheader('🗂️ 請上傳CSV檔案')


