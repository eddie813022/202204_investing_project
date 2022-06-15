### run.py

GUI介面提供選項爬取investing.com的資料，並儲存成EXCEL格式。

主要使用`requests`套件進行get&post的請求來獲得網站上的資料，

並利用`BeautifulSoup`、`pandas.readhtml()`讀取資料格式後，

使用`openpyxl`將數據資料篩選出有用資訊並儲存至EXCEL中，

使用者圖形化介面使用`tkinter`來編寫。

![image](https://github.com/eddie813022/202204_investing_project/blob/main/IMG/run.png)

#### 抓取範例

![image](https://github.com/eddie813022/202204_investing_project/blob/main/IMG/example_excel.png)

### datarun.py

GUI介面提供上傳EXCEL自定義的SHEET公式，合併到原有的EXCEL上產生新的EXCEL(保留格式)。

![image](https://github.com/eddie813022/202204_investing_project/blob/main/IMG/datarun.png)
