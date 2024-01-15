import streamlit as st

def dropdown():
	option = st.sidebar.selectbox('系統說明',['請選擇','操作說明','sender.ini','mail_service.ini' ,'products.csv'])
	if option == '操作說明':
		st.subheader('操作說明')
		st.markdown("""---""")
		st.info('* 上傳商城平台訂單CSV檔 \n * 訂單起訖日：預設值時間區間是今天日期往前推7天 \n * 請選擇郵寄服務種類 \n * 請選擇要發送的寄件人')
		st.error('重要：因為郵局Excel 範例檔案包含程式，必須要在下載檔案後，\
			使用Excel 應用程式開啟，Excel 檔中的程式被執行之後，儲存Excel 檔案，最後才能將檔案上傳到郵局系統，\
			否則郵局系統會發生錯誤')
	elif option == 'sender.ini':
		st.subheader('設定檔sender.ini')
		st.markdown("""---""")
		st.info('* 請在寄件人設定中輸入送件人資訊 \n * 可以設定多筆寄件人，設定規則是在中括弧中輸入寄件人姓名或暱稱(請使用英文)，下方則輸入寄件人資訊 \
			\n * 系統會自動抓取中括弧的寄件人名稱，顯示在介面上的寄件人下拉選單')
		st.markdown('[更新設定檔網址](https://github.com/Ryan-Fukuoka/ship/blob/main/sender.ini)')
	elif option == 'mail_service.ini':
		st.subheader('設定檔mail_service.ini')
		st.markdown("""---""")
		st.info('* 此設定用做輸入Excel 中「郵件種類」欄位，分成quick(國際快捷), package(國際包裹), normal(國際平常小包) 三個設定 \
			\n * package(國際包裹) 多一個reject 設定，對應到Excel 中的欄位「無法投遞處理方式」')
		st.markdown('[更新設定檔網址](https://github.com/Ryan-Fukuoka/ship/blob/main/mail_service.ini)')
	elif option == 'country.ini':
		st.subheader('設定檔country.ini')
		st.markdown("""---""")
		st.info('* 因為商城與郵局國家命名規則不同，\
			請在country 設定兩份文件的國別，設定規則：商城文件中「Ship Country」欄位 = 郵局文件中「收件人國別」，\
			如果沒有設定，將顯示空白，例：United States=U.S.A.')
		st.markdown('[更新設定檔網址](https://github.com/Ryan-Fukuoka/ship/blob/main/country.ini)')
	elif option == 'products.csv':
		st.subheader('產品列表products.csv')
		st.info('* 請在CSV 檔中輸入產品資訊，欄位名稱依序是Item Name,length,width,high,weight,\
			content,currency,description,price\
			\n * 商城文件中的Item Name 必須要和products.csv 中的Item Name 相符，否則會無法對應到產品資訊\
			\n * 使用Excel 修改CSV 檔可能會造成中文編碼錯誤，請儘量使用純文字檔編輯器或其他軟體修改')
		st.markdown('[更新產品列表網址](https://github.com/Ryan-Fukuoka/ship/blob/main/products.csv)')
