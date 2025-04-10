# 全權委託帳戶每日淨值自動化爬蟲工具

## 📌 專案簡介
本專案是一套自動化工具，能分別每日定時爬取六家公司各自所屬的全權委託帳戶產品淨值（含除息資訊），並將資料整理後備份為 Excel 檔案
六家公司對應六個腳本
|公司|腳本檔案|
|---|---|
|中國人壽|china.py|
|台灣人壽|taiwan.py|
|安達人壽|chubb.py|
|法國巴黎人壽|fr_pa.py|
|國泰人壽|cathay.py|
|安聯人壽|allianz.py|

程式使用 Python 撰寫，整合 Selenium 進行網站互動、自動填寫與資料擷取，搭配 `schedule` 套件進行每日定時排程。

## 🎯 功能亮點
- 自動讀取設定的產品名稱與網址，批次爬取各商品歷史淨值資料
- 同步取得「除息金額」，並計算累計除息
- 各家結果分別自動儲存至 Excel，依商品分頁
- 自動建立每日備份資料夾，防止歷史資料遺失
- 可搭配 Windows 工作排程或常駐背景執行，每日定時執行

## 🧩 使用技術
- Python 3
- Selenium + Edge WebDriver
- pandas / numpy / openpyxl
- schedule (排程)
- Excel 資料輸出與備份管理

## 🗂️ 檔案結構
- `allianz.py`：主程式，執行爬蟲與儲存邏輯
- `網址.xlsx`：存放各個帳戶商品名稱與其對應網址（含撥回網址）
- `固定時間.txt`：設定每天自動執行的時間（格式為 hh:mm）
- `excel/`：每日資料儲存主資料夾
- `備份/`：每週備份資料，自動以星期幾編號輪替儲存

## 🔧 執行方式
1. 安裝相依套件：
pip install pandas numpy schedule selenium openpyxl python-dateutil
2. 安裝 Edge WebDriver 並放入與程式相同路徑
3. 確保存在 `網址.xlsx` 檔案並填寫商品資訊
4. 執行各家公司對應、副檔名為.py的程式碼，依據 `固定時間.txt` 內容定時爬取
