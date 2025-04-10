# In[1]:

#公司名稱:安達
name = '安達'
print(name)

# In[2]:
import pandas as pd  #用於處理dataframe的套件
import numpy as np
import schedule  #用於定時的套件
from selenium import webdriver  #用於爬蟲用的套件
#selenium的選項
from selenium.webdriver.edge.options import Options

import time  #時間套件
from datetime import date  #日期套件
from datetime import datetime  #日期套件
from dateutil.relativedelta import relativedelta  #計算日期用的套件
import os  #系統指令套件
import shutil  #檔案處理套件

# In[3]:

def fun_1():    #定義爬蟲用的function

    #建立名位excel的資料夾，用於存放爬蟲結果。
    if os.path.isdir('excel'): #確認備份資料夾是否存在
        print('excel資料夾已存在')
    else:
        os.mkdir('excel')     #沒有備份資料夾的話就建立資料夾
        
    excel_file_path = 'excel/'+name+'_data.xlsx'  #備份路徑與檔名
    print('excel路徑與檔名:' + excel_file_path)
    print('-'*100)


    #備份前一天所抓的檔案
    weekday = str(date.today().weekday()) #以星期作為備份序號
    
    if os.path.isdir('備份'): #確認備份資料夾是否存在
        print('備份資料夾已存在')
    else:
        os.mkdir('備份')     #沒有備份資料夾的話就建立資料夾
    
    file_path = '備份/'+name+'_data_' + weekday + '.xlsx'  #備份路徑與檔名
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
        pass #如果安達_data.xlsx不存在，直接跳過備份
    
    print('-'*100)
    
    
    # 讀取網址的xlsx檔案
    print(name)
    while True: #用while迴圈
        try:    #嘗試讀取網址.xlsx中的工作表
            url_dict = pd.read_excel('網址.xlsx',sheet_name=name,keep_default_na=False)  #keep_default_na=False空值為''
            break #當成功讀取則跳出迴圈
        except:
            print('找不到網址.xlsx')
            print('請確認資料夾內有"網址.xlsx"的檔案或是已關閉Excel應用程式')
    
    
    #開始爬蟲
    options = Options()  
    driver = webdriver.Edge(options=options)
    driver.implicitly_wait(3)  #瀏覽器打開前的等待 
    
    #以for迴圈讀取"網址.xlsx"檔案中的產品名稱與產品網址
    for row in range (url_dict.shape[0]): 
        product = url_dict.iloc[row,0]  #產品名稱
        url = url_dict.iloc[row,1]  #產品網址
        ex_url = url_dict.iloc[row,2]   #產品撥回網址
        ex_format = url_dict.iloc[row,3]
        
        print(product,end=':')
        print(url)
        
        price_table = pd.DataFrame()
        
        #網頁資料讀取
        while price_table.shape[0] <= 5 or price_table.shape[1] <= 1 : #用while迴圈嘗試讀取網頁，如果rows<=5或是columns<=1則重來
            #用while迴圈嘗試讀取網頁，如果讀取到的rows<=5或是columns<=1則重來

            #商品淨值
            try:            
                driver.get(url) #打開產品網址
                time.sleep(3)
                five_years_ago = datetime.now() - relativedelta(years=5)    #計算五年前的日期
                five_years_ago = five_years_ago.strftime('%Y/%m/%d')    #轉換日期的格式
                
                date_input = driver.find_element('id','dFundPriceAction_xDate1')    #找到輸入開始日期的元素
                date_input.clear()  #將開始日期清空
                time.sleep(1)
                date_input.send_keys(five_years_ago)    #輸入五年前的日期
                #time.sleep(5)

                #找到查詢鍵的元素
                serch_button = driver.find_element('id','dFundPriceAction_method:getVAFundPriceHistory')
                #按下查詢鍵
                driver.execute_script("arguments[0].click();", serch_button)
                
                html = driver.page_source #獲取網頁的html
    
                price_table = pd.read_html(html)[8] #找到所要的表格
    
                price_table.columns = price_table.iloc[0]    #將欄位名稱設定為第一列
                price_table = price_table[1:]                #由於資料的第一列仍為名稱，因此資料是從第二列開始到結束
                price_table.columns = ['日期'] + list(price_table.columns[1::])  #將欄位第一個欄位名稱設定為日期
                price_table.set_index('日期',drop=True,inplace=True)    #將日起指定為index
                price_table = price_table.astype(float)  #將資料內容轉換成浮點數
                price_table.rename(columns={'贖回淨值':'漲跌'},inplace=True) #將第二個欄位名稱更改為漲跌
                price_table['漲跌'] = price_table['申購淨值'] - price_table['申購淨值'].shift(-1) #計算漲跌  
                
                
            except Exception as error:  #如果讀取不到資料，則輸出錯誤資訊
                print('網頁讀取失敗，重新嘗試中')
                print("An error occurred:", error)
                print("An exception occurred:", type(error).__name__)
                time.sleep(5)
                
            if ex_url != '':
                try:
                    driver.get(ex_url)  #打開撥回的網址
                    time.sleep(2)
                    ex_html = driver.page_source #獲取網頁的html
                    ex_Dividend = pd.read_html(ex_html)[0]  #找到所需要的表格
                    ex_Dividend.columns = list(ex_Dividend.iloc[2,:]) #將表格的column更改為正確名稱
                    ex_Dividend = ex_Dividend.iloc[3::,]    #保留需要的]資料
                    #更改column名稱，並以日期為索引
                    ex_Dividend = ex_Dividend.rename(columns={'提減 (撥回)日':'日期'}).set_index('日期')
                    ex_Dividend[ex_format] = ex_Dividend[ex_format].apply(pd.to_numeric) #轉換格式為數字
                    ex_Dividend = ex_Dividend.groupby('日期',).sum(numeric_only=True)

                    #ex_Dividend.to_excel('div.xlsx')
                    
                    price_table['除息'] = ex_Dividend[ex_format]
                    price_table['除息'] = price_table['除息'].fillna(0)     #將除息的nan值補零  
                    price_table.sort_values(by='日期',ascending=True,inplace=True)  #將日期升冪，以利計算累計除息
                    price_table['累計除息'] = price_table['除息'].cumsum()  #計算累計除息
                    price_table.sort_values(by='日期',ascending=False,inplace=True) #將日期降冪
                
                except Exception as error:  #如果讀取不到資料，則輸出錯誤資訊
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
            price_table.to_excel(excel_file_path  , sheet_name=product)
            pass
        
        time.sleep(3)
    
            
    driver.close()   #關閉瀏覽器
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