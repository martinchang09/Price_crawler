# In[1]:

#公司名稱:國泰
name = '國泰'
print(name)

# In[2]:
import pandas as pd  #用於處理dataframe的套件
import numpy as np 
import schedule  #用於定時的套件
from selenium import webdriver  #用於爬蟲用的套件

#Edge的選項
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

import time  #時間套件
from datetime import date  #日期套件
from datetime import datetime  #日期套件
import os  #系統指令套件
import shutil  #檔案處理套件
import json  #讀取json的套件

# In[3]:

def fun_1():  #定義爬蟲用的function
	#定義處理json的funtion，用以找出所需資料的網址
	def process_browser_log_entry(entry):
				response = json.loads(entry['message'])['message']
				return response

    
	#建立名位excel的資料夾，用於存放爬蟲結果。
	if os.path.isdir('excel'): #確認備份資料夾是否存在
			print('excel資料夾已存在')
	else:
			os.mkdir('excel')     #沒有備份資料夾的話就建立資料夾
			
	excel_file_path = 'excel/'+name+'_data.xlsx'  #備份路徑與檔名
	print('excel路徑與檔名:' + excel_file_path) 
	print('-'*100)
        

  	#備份前一天所抓的檔案
	week_day = str(date.today().weekday()) #以星期作為備份序號
	
	if os.path.isdir('備份'): #確認備份資料夾是否存在
			print('備份資料夾已存在')
	else:
			os.mkdir('備份')     #沒有備份資料夾的話就建立資料夾
	
	file_path = '備份/'+name+'_data_' + week_day + '.xlsx'  #備份路徑與檔名
	print('備份路徑與檔名:' + file_path)

	try:
		if os.path.isfile(file_path):   #確認備份檔案是否已存在
			os.remove(file_path)        #若存在則刪除
			print('已刪除一周前的備份')
				
			shutil.copy2(excel_file_path,file_path)  #刪除後重新備份
			print('備份成功')
		else:
			shutil.copy2(excel_file_path,file_path)  #如果備份檔案不存在則直接備份檔案
			print('備份成功')
	except:
			pass #如果國泰_data.xlsx不存在，直接跳過備份
	
	print('-'*100)
	

	# 讀取網址的xlsx檔案
	print(name)
	while True: #用while迴圈
		try:    #嘗試讀取網址.xlsx中的國泰工作表
			url_dict = pd.read_excel('網址.xlsx',sheet_name=name,keep_default_na=False)  #keep_default_na=False空值為''
			break #當成功讀取則跳出迴圈
		except:
			print('找不到網址.xlsx')
			print('請確認資料夾內有"網址.xlsx"的檔案或是已關閉Excel應用程式')
	
	
	#開始爬蟲
	options = Options()	
	caps = DesiredCapabilities.EDGE
	caps['ms:loggingPrefs'] = {'performance': 'ALL'}
	driver = webdriver.Edge(capabilities=caps,options=options)
	driver.implicitly_wait(3)  #瀏覽器打開前的等待 
	
	#以for迴圈讀取"網址.xlsx"檔案中的產品名稱與產品網址
	for row in range (url_dict.shape[0]): 
		product = url_dict.iloc[row,0] #產品名稱
		url = url_dict.iloc[row,1]	#產品網址
		ex_url = url_dict.iloc[row,2]	#產品撥回的網址	
		
		ex_format = url_dict.iloc[row,3]		
		
		print(product,end=':')
		print(url)
		
		price_table = pd.DataFrame()
		
		#網頁資料讀取
		while price_table.shape[0] <= 5 or price_table.shape[1] <= 1 : #用while迴圈嘗試讀取網頁，如果rows<=5或是columns<=1則重來
			try:            
					
				driver.get(url)  #打開產品網址
				time.sleep(3)
				browser_log = []
				browser_log = driver.get_log('performance')	#取得網頁的log
				
				#用前面定義的funtion處理log
				events = [process_browser_log_entry(entry) for entry in browser_log]
				#找到log中需要的目標
				events = [event for event in events if 'Network.response' in event['method']]
				
				u = ''
				
				for i in range(len(events)):
					for j,k in events[i].items():
						if str(k).find('djbcd') != -1:  #找到djbcd
							#print(product,end=':')
							#找到網址
							u = events[i]['params']['response']['url']
							#print(u)
				
				driver.get(u)  #打開抓到的網址
				print('正在抓'+u)
				time.sleep(3)

				#找到所需的元素，並轉成文字
				price_list = driver.find_element('tag name',"body").text
				
				#資料前半是日期
				price_date = price_list.split(' ')[0].split(',')
				#資料後半是價格
				price = price_list.split(' ')[1].split(',')
				
		
				price_table['日期'] = price_date
				price_table['日期'] = pd.to_datetime(price_table['日期']).dt.strftime('%Y-%m-%d')
				price_table['淨值'] = price
				price_table['淨值'] = price_table['淨值'].apply(float)  
				price_table['漲跌'] = price_table['淨值'] - price_table['淨值'].shift(-1)  #計算漲跌
				price_table['漲跌幅(%)'] = (price_table['淨值'] / price_table['淨值'].shift(-1)-1)*100 #計算漲跌幅
				price_table.sort_values(by='日期',ascending=False, inplace = True)
				price_table.set_index('日期',inplace=True) #以日期為索引
				
				url = ''
				events = []
						

			except Exception as error:	#如果讀取不到資料，則輸出錯誤資訊
				print('網頁讀取失敗，重新嘗試中')
				print("An error occurred:", error)
				print("An exception occurred:", type(error).__name__)
				time.sleep(5)
			
			if ex_url != '':
                #商品撥回
				try:
					driver.get(ex_url) #打開撥回的網址
					time.sleep(2)
					ex_html = driver.page_source  #獲取網頁的html
					ex_Dividend = pd.read_html(ex_html)[0]  #用pandas讀取html中的表格
					#更改column的名稱，並以日期為索引
					ex_Dividend['除息日'] = pd.to_datetime(ex_Dividend['除息日']).dt.strftime('%Y-%m-%d') #統一日期格式
					ex_Dividend = ex_Dividend.rename(columns={'除息日':'日期'}).set_index('日期')
					ex_Dividend = ex_Dividend.groupby('日期',).sum(numeric_only=True)
					#ex_Dividend.to_excel(product+'.xlsx')

					#ex_Dividend.to_excel('div.xlsx')
					price_table['除息'] = ex_Dividend[ex_format]
					price_table['除息'] = price_table['除息'].fillna(0)   #將除息的nan值補零
					price_table.sort_values(by='日期',ascending=True,inplace=True)	#將日期升冪，以利計算累計除息
					price_table['累計除息'] = price_table['除息'].cumsum()	#計算累計除息
					price_table.sort_values(by='日期',ascending=False,inplace=True)	#將日期降冪
						
				except Exception as error:	#如果讀取不到資料，則輸出錯誤資訊
					print('撥回網頁讀取失敗，重新嘗試中')
					print("An error occurred:", error)
					print("An exception occurred:", type(error).__name__)
					time.sleep(5)

			else:
				#如果沒有除息的網址則為0
				price_table['除息'] = 0
				price_table['累計除息'] = 0
				
		#資料寫入EXCEL
		try:
			 #將excel檔打開，把資料寫入該商品名稱的工作表
			with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:
				price_table.to_excel(writer, sheet_name=product)   
		except:
            #如果打不開檔案，則直接存檔(檔案會僅有一張工作表)
			price_table.to_excel(excel_file_path, sheet_name=product)
			pass
	
		time.sleep(3)

					
	driver.close()	#關閉瀏覽器
	print(name+'已完成')
	print('-'*100)


# In[ ]:
try:
    #打開"固定時間.txt",讀取檔案內所寫的時間
    run_time = open('固定時間.txt','r').read()
except:
    #如果沒有檔案則手動輸入
    run_time = input('請輸入指定的時間(格式為hh:mm):')

print('每日運行時間:'+run_time)

#設定上面定義的funtion在每天指定的時間運作
schedule.every().day.at(run_time).do(fun_1)

print('如果要停止程式，請按"ctrl"+"c"!')

#以while迴圈執行
while True:

    schedule.run_pending()

    time.sleep(1)