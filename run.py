from crawler_request import get_pathconfig,get_exchang_dict,create_stockdata,get_stock_list,get_country_dict,read_stock_list,get_dateconfig,write_dateconfig,get_keypoint_stocklist,get_checkboxconfig,write_pathconfig
from tkinter import ttk,filedialog
from tkinter.scrolledtext import ScrolledText
from tkcalendar import DateEntry
from pathlib import Path
from openpyxl import load_workbook
from icon import img
import babel.numbers
import tkinter as tk
import threading
import base64
import time
import os

# path configure
path = Path.cwd()

"""gui configure"""
tmp = open("tmp.ico","wb+")
tmp.write(base64.b64decode(img))
tmp.close()
window = tk.Tk()
# gui favicon configure
window.iconbitmap("tmp.ico")
window.title('investing_v1.6')
window.resizable(False,False)
# gui theme configure
windw_style = ttk.Style(window)
windw_style.theme_use('clam')
# progressbar configure
windw_style.configure("red.Horizontal.TProgressbar", foreground='green', background='green')
os.remove("tmp.ico")

"""define varriable"""
# country_list = [""]
stop_threads = False
country_dict = get_country_dict()
country_list = [ i for i in country_dict.keys() ]
country_list.insert(0,"")
search_list= []
market_list = []
category_list = [{'所有板塊': '-1'}, {'消费类（非周期性）': '8'}, {'交通运输': '10'}, {'公用事业': '22'},
                 {'医疗保健': '18'}, {'基础材料': '7'}, {'工業': '15'}, {'房地產': '23'}, {'服务': '2'},
                 {'材料': '14'}, {'消費者日常用品': '17'}, {'消费类（周期性）': '3'}, {'生产资料': '5'},
                 {'科技': '4'}, {'综合性企业': '12'}, {'能源': '13'}, {'資訊科技': '20'}, {'通訊服務': '21'},
                 {'金融': '19'}, {'非必需消費品': '16'}]
frame1_1_inside_text_list = ["總收入","收入","其他收入合計","稅收成本合計","毛利","經營開支總額","銷售/一般/管理費用合計",
                             "研發","折舊/攤銷","利息開支（收入）- 營運淨額","例外開支（收入）","其他運營開支總額","營業收入",
                             "利息收入（開支）- 非營運淨額","出售資產收入（虧損）","其他，淨額","稅前淨收益","備付所得稅",
                             "稅後淨收益","少數股東權益","附屬公司權益","美國公認會計準則調整","計算特殊項目前的淨收益","特殊項目合計",
                             "淨收入","淨收入調整總額","扣除特殊項目的普通收入","稀釋調整","稀釋后淨收入","稀釋后加權平均股",
                             "稀釋后扣除特殊項目的每股盈利","每股股利 – 普通股首次發行","稀釋后每股標準盈利"]
frame1_2_inside_text_list = ["總收入","保費收入合計","投資收益淨額","變現收益（虧損）","其他收入合計","經營開支總額","虧損、福利和修訂合計",
                             "购置成本攤銷","銷售/一般/管理費用合計","折舊/攤銷","利息開支（收入）- 營運淨額","例外開支（收入）","其他運營開支總額",
                             "營業收入","利息收入（開支）- 非營運淨額","出售資產收入（虧損）","其他，淨額","稅前淨收益","備付所得稅",
                             "稅後淨收益","少數股東權益","附屬公司權益","美國公認會計準則調整","計算特殊項目前的淨收益","特殊項目合計",
                             "淨收入","淨收入調整總額","扣除特殊項目的普通收入","稀釋調整","稀釋后淨收入","稀釋后加權平均股",
                             "稀釋后扣除特殊項目的每股盈利","每股股利 – 普通股首次發行","稀釋后每股標準盈利"]
frame1_3_inside_text_list = ["利息收益淨額","銀行利息收入","利息開支總額","風險準備金","扣除風險準備金後淨利息收入","銀行非利息收入",
                             "銀行非利息開支","稅前淨收益","備付所得稅","稅後淨收益","少數股東權益","附屬公司權益","美國公認會計準則調整",
                             "計算特殊項目前的淨收益","特殊項目合計","淨收入","淨收入調整總額","扣除特殊項目的普通收入","稀釋調整",
                             "稀釋后淨收入","稀釋后加權平均股","稀釋后扣除特殊項目的每股盈利","每股股利 – 普通股首次發行","稀釋后每股標準盈利"]
frame2_1_inside_text_list = ["流動資產合計","現金和短期投資","現金","現金和現金等價物","短期投資","淨應收款合計","淨交易應收款合計",
                           "庫存合計","預付費用","其他流動資產合計","總資產","物業/廠房/設備淨總額","物業/廠房/設備總額",
                           "累計折舊合計","商譽淨額","無形資產淨額","長期投資","長期應收票據","其他長期資產合計","其他資產合計",
                           "總流動負債","應付賬款","應付/應計","應計費用","應付票據/短期債務","長期負債當前應收部分/資本租賃",
                           "其他流動負債合計","總負債","長期債務合計","長期債務","資本租賃債務","遞延所得稅","少數股東權益",
                           "其他負債合計","總權益","可贖回優先股合計","不可贖回優先股淨額","普通股合計","附加資本","保留盈餘(累計虧損)",
                           "普通庫存股","員工持股計劃債務擔保","未實現收益（虧損）","其他權益合計","負債及股東權益總計",
                           "已發行普通股合計","已發行優先股合計"]
frame2_2_inside_text_list = ["流動資產合計","總資產","現金","現金和現金等價物","淨應收款合計","預付費用","物業/廠房/設備淨總額","物業/廠房/設備總額","累計折舊合計",
                            "商譽淨額","無形資產淨額","長期投資","應收保險","長期應收票據","其他長期資產合計","遞延保單獲得成本","其他資產合計","總流動負債",
                            "總負債","應付賬款","應付/應計","應計費用","保單負債","應付票據/短期債務","長期負債當前應收部分/資本租賃","其他流動負債合計",
                            "長期債務合計","長期債務","資本租賃債務","遞延所得稅","少數股東權益","其他負債合計","總權益","可贖回優先股合計","不可贖回優先股淨額",
                            "普通股合計","附加資本","保留盈餘(累計虧損)","普通庫存股","員工持股計劃債務擔保","未實現收益（虧損）","其他權益合計",
                            "負債及股東權益總計","已發行普通股合計","已發行優先股合計"]
frame2_3_inside_text_list = ["流動資產合計","總資產","銀行應付現金和欠款","其他盈利資產合計","淨貸款","物業/廠房/設備淨總額","物業/廠房/設備總額","累計折舊合計",
                             "商譽淨額","無形資產淨額","長期投資","其他長期資產合計","其他資產合計","總流動負債","總負債","應付賬款","應付/應計","應計費用",
                             "存款總額","其他付息負債合計","短期借貸總額","長期負債當前應收部分/資本租賃","其他流動負債合計","長期債務合計","長期債務","資本租賃債務",
                             "遞延所得稅","少數股東權益","其他負債合計","總權益","可贖回優先股合計","不可贖回優先股淨額","普通股合計","附加資本","保留盈餘(累計虧損)",
                             "普通庫存股","員工持股計劃債務擔保","未實現收益（虧損）","其他權益合計","負債及股東權益總計","已發行普通股合計","已發行優先股合計"]
frame3_1_inside_text_list = ["淨收益/起點","來自經營活動的現金","折舊/遞耗","攤銷","遞延稅","非現金項目","現金收入","現金支出","現金稅金支出",
                           "現金利息支出","營運資金變動","來自投資活動的現金","資本支出","其他投資現金流項目合計","來自融資活動的現金",
                           "融資現金流項目","發放現金紅利合計","股票發行（贖回）淨額","債務發行（贖回）淨額","外匯影響","現金變動淨額",
                           "期初現金結餘","期末現金結餘","自由現金流","自由現金流增長","自由現金流收益率"]

# 建立複選框變數清單(代替range(number))
f1_1_var_namelist = [ "f1_in_f1_va"+str(x+1) for x in range(0,33) ]
f1_2_var_namelist = [ "f1_in_f2_va"+str(x+1) for x in range(0,34) ]
f1_3_var_namelist = [ "f1_in_f3_va"+str(x+1) for x in range(0,24) ]
f2_1_var_namelist = [ "f2_in_f1_va"+str(x+1) for x in range(0,47) ]
f2_2_var_namelist = [ "f2_in_f2_va"+str(x+1) for x in range(0,45) ]
f2_3_var_namelist = [ "f2_in_f3_va"+str(x+1) for x in range(0,42) ]
f3_1_var_namelist = [ "f3_in_f1_va"+str(x+1) for x in range(0,26) ]

# 建立複選框變數清單
frame1_1_check_btn_list = [ "f1_1_check_btn" +str(x+1) for x in range(0,33) ]
frame1_2_check_btn_list = [ "f1_2_check_btn" +str(x+1) for x in range(0,34) ]
frame1_3_check_btn_list = [ "f1_3_check_btn" +str(x+1) for x in range(0,24) ]
frame2_1_check_btn_list = [ "f2_1_check_btn" +str(x+1) for x in range(0,47) ]
frame2_2_check_btn_list = [ "f2_2_check_btn" +str(x+1) for x in range(0,45) ]
frame2_3_check_btn_list = [ "f2_3_check_btn" +str(x+1) for x in range(0,42) ]
frame3_1_check_btn_list = [ "f3_1_check_btn" +str(x+1) for x in range(0,26) ]

# event functions
def country_select_event():
    global market_list
    country_name = country_btn.get()
    if country_name:
        if country_name in country_list:
            country_code = country_dict[country_name]
            stext.config(state="normal")
            stext.insert(tk.END,"=====================================\n")
            stext.insert(tk.END,"正在查詢"+country_name+"交易所數量..\n")
            exchange_dict = get_exchang_dict({country_name:country_code})
            market_btn.set("")        # 將選項清空
            market_list = exchange_dict
            market_name_list = [ i for i in market_list ]
            market_btn["value"] = market_name_list
            market_btn.current(0)
            category_name_list = [ j for i in category_list for j,k in i.items() ]
            category_btn["value"] = category_name_list
            category_btn.current(0)
            stext.insert(tk.END,country_name+"共有"+str(len(market_list)-1)+"筆交易所.\n")
            stext.insert(tk.END,"=====================================\n")
            stext.see(tk.END)
            stext.config(state="disable")
        else:
            stext.config(state="normal")
            stext.insert(tk.END,"未正確選擇國家\n")
            stext.insert(tk.END,"=====================================\n")
            stext.see(tk.END)
            stext.config(state="disable")

def country_enter_envet():
    global search_list
    if search_list:
        search_list.clear()
    word = country_btn.get()
    search_list = get_keypoint_stocklist(word)
    if search_list:
        for i in search_list:
            stext.config(state="normal")
            stext.insert(tk.END,i[0]+"\n")
            if i != search_list[-1]:
                stext.insert(tk.END,"-------------------------------------\n")
            stext.see(tk.END)
        stext.insert(tk.END,"=====================================\n")
        stext.config(state="disable")
        
def country_btn_event(event):
    country_btn_td = threading.Thread(target=country_select_event)
    country_btn_td.setDaemon(True)
    country_btn_td.start()
    
def country_btn_enter_event(event):
    country_select_td = threading.Thread(target=country_enter_envet)
    country_select_td.setDaemon(True)
    country_select_td.start()

def loading_evnet(num):
    progress["value"] = 0
    if num == 16:
        timer = 6.5
    elif num == 14:
        timer = 7
    for i in range(num):
        if progress["value"] < progress["maximum"]:
            progress["value"] += timer
            window.update()
            time.sleep(1)
    progress["value"] = 100
    window.update()
    
def loading2_evnet(num):
    if num < 100:
        addnum = int(100 / num)
        while progress["value"] < progress["maximum"]:
            progress["value"] += addnum
            window.update()
            time.sleep(16)
            if stop_threads:
                progress["value"] = 100
                break
    elif num > 100:
        addnum = int(num / 100)
        while progress["value"] < progress["maximum"]:
            progress["value"] += 1
            window.update()
            time.sleep(addnum*16)
            if stop_threads:
                progress["value"] = 100
                break
    progress["value"] = 100
    window.update()

def start_date_btn_event(event):
    date_ = start_date_btn.get()
    write_dateconfig(date_)

def singlecrawler_btn_event():
    single_btn_td = threading.Thread(target=singledownload_event)
    single_btn_td.setDaemon(True)
    single_btn_td.start()
    
def bulkcrawler_btn_event():
    bulk_btn_td = threading.Thread(target=bulkdownload_event)
    bulk_btn_td.setDaemon(True)
    bulk_btn_td.start()

def crawlerstocklist_btn_event():
    stocklist_btn_td = threading.Thread(target=crawlerstocklist_event)
    stocklist_btn_td.setDaemon(True)
    stocklist_btn_td.start()
    
def stopcrawler_btn_event():
    stop_btn_td =threading.Thread(target=stopcrawler_event)
    stop_btn_td.setDaemon(True)
    stop_btn_td.start()
    
def crawlerstocklist_event():
    progress["value"] = 0
    stocktimer_lb.place_forget() 
    crawlerstocklist_btn.config(state="disable")
    bulkcrawler_btn.config(state="disable")
    bulk_start_btn.config(state="disable")
    singlecrawler_btn.config(state="disable")
    download_path = Path(downloadpath_text.get("1.0","end-1c"))
    country_name = country_btn.get()
    exchange_name = market_btn.get()
    category_name = category_btn.get()
    prefilename1 =  country_name+"-"+exchange_name+"-"+category_name+".xlsx"
    prefilename2 = country_name+"-"+exchange_name+".xlsx"
    prefilename3 = country_name+".xlsx"
    if exchange_name and category_name:
        save_name = download_path / prefilename1
    elif exchange_name:
        save_name = download_path / prefilename2
    else:
        save_name = download_path / prefilename3
    try:
        country_code = country_dict[country_name]
    except:
        stext.config(state="normal")
        stext.insert(tk.END,"=====================================\n")
        stext.insert(tk.END,"請先選取國家及交易所.\n")
        stext.insert(tk.END,"=====================================\n")
        stext.see(tk.END)
        stext.config(state="disable")
        country_code = ""
    if market_list:
        try:
            exchange_code = market_list.get(exchange_name)
            if exchange_code == "-1":
                exchange_code = "a"
        except:
            exchange_code = "a"
    else:
        exchange_code = "a"
    try:
        for i in category_list:
            if category_name in i.keys():
                category_code = i.get(category_name)
                if category_code == "-1":
                    category_code = "a"
    except:
        category_code = "a"
    if country_code and exchange_code and category_code:
        stock_data = {"cname":country_name,
                    "ccode":country_code,
                    "ename":exchange_name,
                    "ecode":exchange_code,
                    "gname":category_name,
                    "gcode":category_code,
                    "progress":progress,
                    "save":save_name}
        stext.config(state="normal")
        stext.insert(tk.END,"=====================================\n")
        stext.insert(tk.END,f"正在進行{country_name}-{exchange_name}-{category_name}個股清單抓取.\n")
        stext.see(tk.END)
        stext.config(state="disable")
        stock_totalcount = str(get_stock_list(**stock_data))
        if int(stock_totalcount) <1:
            stext.config(state="normal")
            stext.insert(tk.END,f"{country_name}-{exchange_name}-{category_name}查無個股資料.\n")
            stext.insert(tk.END,"=====================================\n")
            stext.see(tk.END)
            stext.config(state="disable")
        else:
            stext.config(state="normal")
            stext.insert(tk.END,f"完成{country_name}-{exchange_name}-{category_name}個股清單抓取({stock_totalcount}).\n")
            stext.insert(tk.END,"=====================================\n")
            stext.see(tk.END)
            stext.config(state="disable")
    else:
        progress["value"] = 100
    crawlerstocklist_btn.config(state="normal")
    bulkcrawler_btn.config(state="normal")
    bulk_start_btn.config(state="normal")
    singlecrawler_btn.config(state="normal")
     
def singledownload_event():
    stocktimer_lb.place_forget() 
    progress["value"] = 0
    entertext = ""
    crawlerstocklist_btn.config(state="disable")
    bulkcrawler_btn.config(state="disable")
    bulk_start_btn.config(state="disable")
    singlecrawler_btn.config(state="disable")
    start = time.perf_counter()
    tempfile = Path.cwd() / "temp.xlsx"
    checkwb = load_workbook(tempfile)
    checkws = checkwb.active
    try:
        f1general_choose_list = [ globals()["f1_in_f1_va"+str(i)].get() for i in range(len(f1_1_var_namelist)) ]
        f1finance_choose_list = [ globals()["f1_in_f2_va"+str(i)].get() for i in range(len(f1_2_var_namelist)) ]
        f1finance2_choose_list = [ globals()["f1_in_f3_va"+str(i)].get() for i in range(len(f1_3_var_namelist)) ]
        f2general_choose_list = [ globals()["f2_in_f1_va"+str(i)].get() for i in range(len(f2_1_var_namelist)) ]
        f2finance_choose_list = [ globals()["f2_in_f2_va"+str(i)].get() for i in range(len(f2_2_var_namelist)) ]
        f2finance2_choose_list = [ globals()["f2_in_f3_va"+str(i)].get() for i in range(len(f2_3_var_namelist)) ]
        f3general_choose_list = [ globals()["f3_in_f1_va"+str(i)].get() for i in range(len(f3_1_var_namelist)) ]    
        download_path = Path(downloadpath_text.get("1.0","end-1c"))
        start_date = start_date_btn.get()
        end_date = end_date_btn.get()
        for i in range(2,35):
            checkws.cell(row=i,column=2,value=f1general_choose_list[i-2])
        for i in range(35,69):
            checkws.cell(row=i,column=2,value=f1finance_choose_list[i-35])
        for i in range(69,93):
            checkws.cell(row=i,column=2,value=f1finance2_choose_list[i-69])
        for i in range(93,140):
            checkws.cell(row=i,column=2,value=f2general_choose_list[i-93])
        for i in range(140,185):
            checkws.cell(row=i,column=2,value=f2finance_choose_list[i-140])
        for i in range(185,227):
            checkws.cell(row=i,column=2,value=f2finance2_choose_list[i-185])
        for i in range(227,253):
            checkws.cell(row=i,column=2,value=f3general_choose_list[i-227])
        checkwb.save(tempfile)
        entertext = stext.selection_get()
    except:
        stext.config(state="normal")
        stext.insert(tk.END,"=====================================\n")
        stext.insert(tk.END,"請輸入個股關鍵字或代號查詢並選取\n")
        stext.insert(tk.END,"=====================================\n")
        stext.see(tk.END)
        stext.config(state="disable")
    if entertext:
        if entertext in country_list:
            stext.config(state="normal")
            stext.insert(tk.END,"請先輸入個股關鍵字或代號查詢並選取\n")
            stext.see(tk.END)
            stext.config(state="disable")
        elif entertext in [ i[0] for i in search_list ]:
            progres_td = threading.Thread(target=loading_evnet,args=(16,))
            progres_td.setDaemon(True)
            progres_td.start()
            folder1 = download_path / "一般業"
            folder2 = download_path / "保險業"
            folder3 = download_path / "銀行業"
            folder1.mkdir(exist_ok=True)
            folder2.mkdir(exist_ok=True)
            folder3.mkdir(exist_ok=True)
            for i in search_list:
                if entertext == i[0]:
                    print(i[0])
                    stext.config(state="normal")
                    stext.insert(tk.END,"=====================================\n")
                    stext.insert(tk.END, f"開始抓取({entertext})個股.\n")
                    stext.insert(tk.END,"=====================================\n")
                    stext.see(tk.END)
                    stext.config(state="disable")
                    create_stockdata(i=i,sdate=start_date,edate=end_date,f1_1_list=f1general_choose_list,f1_2_list=f1finance_choose_list,
                                    f1_3_list=f1finance2_choose_list,f2_1_list=f2general_choose_list,f2_2_list=f2finance_choose_list,
                                    f2_3_list=f2finance2_choose_list,f3_1_list=f3general_choose_list,download_path=download_path)
                    stext.config(state="normal")
                    stext.insert(tk.END,"=====================================\n")
                    stext.insert(tk.END,f"完成抓取{entertext}個股.\n")
                    stext.insert(tk.END,"=====================================\n")
                    stext.see(tk.END)
                    stext.config(state="disable")
    else:
        progress["value"] = 100
    print(f"Cost: {time.perf_counter() - start}")
    crawlerstocklist_btn.config(state="normal")
    bulkcrawler_btn.config(state="normal")
    bulk_start_btn.config(state="normal")
    singlecrawler_btn.config(state="normal")
    
def bulkdownload_event():
    global stop_threads
    stop_threads = False
    progress["value"] = 0
    tempfile = Path.cwd() / "temp.xlsx"
    checkwb = load_workbook(tempfile)
    checkws = checkwb.active
    try:
        stock_path = Path(stocklistpath_text.get("1.0","end-1c"))
    except:
        progress["value"] = 100
        stext.config(state="normal")
        stext.insert(tk.END,"=====================================\n")
        stext.insert(tk.END,"請先指定個股清單.\n")
        stext.insert(tk.END,"=====================================\n")
        stext.see(tk.END)
        stext.config(state="disable")
    try:
        wb = load_workbook(stock_path)
        ws = wb.active
        f1general_choose_list = [ globals()["f1_in_f1_va"+str(i)].get() for i in range(len(f1_1_var_namelist)) ]
        f1finance_choose_list = [ globals()["f1_in_f2_va"+str(i)].get() for i in range(len(f1_2_var_namelist)) ]
        f1finance2_choose_list = [ globals()["f1_in_f3_va"+str(i)].get() for i in range(len(f1_3_var_namelist)) ]
        f2general_choose_list = [ globals()["f2_in_f1_va"+str(i)].get() for i in range(len(f2_1_var_namelist)) ]
        f2finance_choose_list = [ globals()["f2_in_f2_va"+str(i)].get() for i in range(len(f2_2_var_namelist)) ]
        f2finance2_choose_list = [ globals()["f2_in_f3_va"+str(i)].get() for i in range(len(f2_3_var_namelist)) ]
        f3general_choose_list = [ globals()["f3_in_f1_va"+str(i)].get() for i in range(len(f3_1_var_namelist)) ]    
        download_path = Path(bulkdownloadpath_text.get("1.0","end-1c"))
        start_date = start_date_btn.get()
        end_date = end_date_btn.get()
        for i in range(2,35):
            checkws.cell(row=i,column=2,value=f1general_choose_list[i-2])
        for i in range(35,69):
            checkws.cell(row=i,column=2,value=f1finance_choose_list[i-35])
        for i in range(69,93):
            checkws.cell(row=i,column=2,value=f1finance2_choose_list[i-69])
        for i in range(93,140):
            checkws.cell(row=i,column=2,value=f2general_choose_list[i-93])
        for i in range(140,185):
            checkws.cell(row=i,column=2,value=f2finance_choose_list[i-140])
        for i in range(185,227):
            checkws.cell(row=i,column=2,value=f2finance2_choose_list[i-185])
        for i in range(227,253):
            checkws.cell(row=i,column=2,value=f3general_choose_list[i-227])
        checkwb.save(tempfile)
        last_index = int(read_stock_list(stock_path))
        if last_index:
            folder1 = download_path / "一般業"
            folder2 = download_path / "保險業"
            folder3 = download_path / "銀行業"
            folder1.mkdir(exist_ok=True)
            folder2.mkdir(exist_ok=True)
            folder3.mkdir(exist_ok=True)
            counter_num = (ws.max_row - last_index)+1
            counter_time = counter_num * 16
            nowtime = time.time() # 時間戳
            finaltime = nowtime + counter_time
            time_struct = time.localtime(finaltime) # 時間元組
            time_string = time.strftime("%Y-%m-%d %H:%M", time_struct) # 字串
            stocktimer_lb.configure(text=time_string)
            stocktimer_lb.place(relx=0.53,rely=0.88)
            stext.config(state="normal")
            stext.insert(tk.END,"=====================================\n")
            stext.insert(tk.END,f"預計完成時間為{time_string}.\n") 
            stext.insert(tk.END,"=====================================\n")
            stext.see(tk.END)
            stext.config(state="disable")
            progres_bk = threading.Thread(target=loading2_evnet,args=(counter_num,))
            progres_bk.setDaemon(True)
            progres_bk.start()
            for i in range(last_index,ws.max_row+1):
                name = ws.cell(row=last_index,column=1).value
                link = ws.cell(row=last_index,column=2).value
                id = ws.cell(row=last_index,column=3).value
                irow = (name,link,id)
                stext.config(state="normal")
                stext.insert(tk.END,"=====================================\n")
                stext.insert(tk.END, f"開始抓取第{last_index}個，({name})個股.\n")
                stext.see(tk.END)
                stext.config(state="disable")
                create_stockdata(i=irow,sdate=start_date,edate=end_date,f1_1_list=f1general_choose_list,f1_2_list=f1finance_choose_list,
                                f1_3_list=f1finance2_choose_list,f2_1_list=f2general_choose_list,f2_2_list=f2finance_choose_list,
                                f2_3_list=f2finance2_choose_list,f3_1_list=f3general_choose_list,download_path=download_path)
                if stop_threads:
                    ws.cell(row=last_index,column=4,value="finished")
                    stext.config(state="normal")
                    stext.insert(tk.END,"=====================================\n")
                    stext.insert(tk.END, f"暫停至第{last_index}個，({name})個股抓取.\n")
                    stext.insert(tk.END,"=====================================\n")
                    stext.see(tk.END)
                    stext.config(state="disable")
                    last_index += 1
                    wb.save(stock_path)
                    stocktimer_lb.place_forget() 
                    break
                ws.cell(row=last_index,column=4,value="finished")
                stext.config(state="normal")
                stext.insert(tk.END, f"完成抓取第{last_index}個，({name})個股.\n")
                stext.insert(tk.END,"=====================================\n")
                stext.see(tk.END)
                stext.config(state="disable")    
                last_index += 1
            wb.save(stock_path)
            progress["value"] = 100
        else:
            stext.config(state="normal")
            stext.insert(tk.END,"=====================================\n")
            stext.insert(tk.END,"該個股清單已抓取完成.\n")
            stext.insert(tk.END,"=====================================\n")
            stext.see(tk.END)
            stext.config(state="disable")
            wb.save(stock_path)
            progress["value"] = 100
    except Exception as e:
        print(e)
        stext.config(state="normal")
        stext.insert(tk.END,"=====================================\n")
        stext.insert(tk.END,"請指定正確個股清單.\n")
        stext.insert(tk.END,"=====================================\n")
        stext.see(tk.END)
        stext.config(state="disable")
        wb.save(stock_path)
    progress["value"] = 100

def stopcrawler_event():
    global stop_threads 
    progress["value"] = 0
    stocktimer_lb.place_forget() 
    if stop_threads:
        stext.config(state="normal")
        stext.insert(tk.END,"=====================================\n")
        stext.insert(tk.END,"目前並無任務執行中.\n")
        stext.insert(tk.END,"=====================================\n")
        stext.see(tk.END)
        stext.config(state="disable")
    stop_threads = True
    progress["value"] = 100
    print("kill successful")

def align_center(root, width, height):
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width)/2, (screenheight - height)/2)
    root.geometry(size)

def f1_f1_selectall_event():
    for i in range(len(frame1_1_check_btn_list)):
        globals()["f1_in_f1_btn"+str(i)].select()

def f1_f1_dselectall_event():
    for i in range(len(frame1_1_check_btn_list)):
        globals()["f1_in_f1_btn"+str(i)].deselect()
        
def f1_f2_selectall_event():
    for i in range(len(frame1_2_check_btn_list)):
        globals()["f1_in_f2_btn"+str(i)].select()

def f1_f2_dselectall_event():
    for i in range(len(frame1_2_check_btn_list)):
        globals()["f1_in_f2_btn"+str(i)].deselect()

def f1_f3_selectall_event():
    for i in range(len(frame1_3_check_btn_list)):
        globals()["f1_in_f3_btn"+str(i)].select()
        
def f1_f3_dselectall_event():
    for i in range(len(frame1_3_check_btn_list)):
        globals()["f1_in_f3_btn"+str(i)].deselect()

def f2_f1_selectall_event():
    for i in range(len(frame2_1_check_btn_list)):
        globals()["f2_in_f1_btn"+str(i)].select()

def f2_f1_dselectall_event():
    for i in range(len(frame2_1_check_btn_list)):
        globals()["f2_in_f1_btn"+str(i)].deselect()
        
def f2_f2_selectall_event():
    for i in range(len(frame2_2_check_btn_list)):
        globals()["f2_in_f2_btn"+str(i)].select()

def f2_f2_dselectall_event():
    for i in range(len(frame2_2_check_btn_list)):
        globals()["f2_in_f2_btn"+str(i)].deselect()
        
def f2_f3_selectall_event():
    for i in range(len(frame2_3_check_btn_list)):
        globals()["f2_in_f3_btn"+str(i)].select()

def f2_f3_dselectall_event():
    for i in range(len(frame2_3_check_btn_list)):
        globals()["f2_in_f3_btn"+str(i)].deselect()  
        
def f3_f1_selectall_event():
    for i in range(len(frame3_1_check_btn_list)):
        globals()["f3_in_f1_btn"+str(i)].select()

def f3_f1_dselectall_event():
    for i in range(len(frame3_1_check_btn_list)):
        globals()["f3_in_f1_btn"+str(i)].deselect()

def samplepath_event():
    filename = filedialog.askopenfilename(parent=window,initialdir="C:/",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
    samplepath_text.config(state="normal")
    samplepath_text.delete(1.0,"end")
    samplepath_text.insert(1.0, filename)
    samplepath_text.config(state="disable")
    
def downloadpath_event():
    filename = filedialog.askdirectory(parent=window)
    if not filename:
        filename = "C:/"
    downloadpath_text.config(state="normal")
    downloadpath_text.delete(1.0,"end")
    downloadpath_text.insert(1.0, filename)
    downloadpath_text.config(state="disable")
    write_pathconfig(type="single",download_path=filename)

def downloadpath2_event():
    filename = filedialog.askdirectory(parent=window)
    if not filename:
        filename = "C:/"
    bulkdownloadpath_text.config(state="normal")
    bulkdownloadpath_text.delete(1.0,"end")
    bulkdownloadpath_text.insert(1.0, filename)
    bulkdownloadpath_text.config(state="disable")
    write_pathconfig(type="bulk",download_path=filename)
    
def stocklistpath_event():
    filename = filedialog.askopenfilename(parent=window,initialdir="C:/",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
    stocklistpath_text.config(state="normal")
    stocklistpath_text.delete(1.0,"end")
    stocklistpath_text.insert(1.0, filename)
    stocklistpath_text.config(state="disable")
    
# 批量建立tkinter變數
def f1_var_name():
    for i in range(len(f1_1_var_namelist)):
        globals()["f1_in_f1_va"+str(i)] = tk.StringVar()
    for i in range(len(f1_2_var_namelist)):
        globals()["f1_in_f2_va"+str(i)] = tk.StringVar()
    for i in range(len(f1_3_var_namelist)):
        globals()["f1_in_f3_va"+str(i)] = tk.StringVar()
        
def f2_var_name():
    for i in range(len(f2_1_var_namelist)):
        globals()["f2_in_f1_va"+str(i)] = tk.StringVar()
    for i in range(len(f2_2_var_namelist)):
        globals()["f2_in_f2_va"+str(i)] = tk.StringVar()
    for i in range(len(f2_3_var_namelist)):
        globals()["f2_in_f3_va"+str(i)] = tk.StringVar()

def f3_var_name():
    for i in range(len(f3_1_var_namelist)):
        globals()["f3_in_f1_va"+str(i)] = tk.StringVar()
        
# 批量建立損益表複選按鈕
def frame1_check_btn():
    key1_list = ["收入","其他收入合計","銷售/一般/管理費用合計","研發","折舊/攤銷","利息開支（收入）- 營運淨額","例外開支（收入）","其他運營開支總額"]
    key2_list = ["保費收入合計","投資收益淨額","變現收益（虧損）","其他收入合計","虧損、福利和修訂合計","购置成本攤銷",
                 "銷售/一般/管理費用合計","折舊/攤銷","利息開支（收入）- 營運淨額","例外開支（收入）","其他運營開支總額"]
    key3_list = ["銀行利息收入","利息開支總額"]
    copy_frame1_1_list = frame1_1_inside_text_list
    copy_frame1_2_list = frame1_2_inside_text_list
    copy_frame1_3_list = frame1_3_inside_text_list
    opt1 = tk.IntVar()
    opt2 = tk.IntVar()
    opt3 = tk.IntVar()
    frame1_inside_1_sall_rbtn = tk.Radiobutton(frame1_inside_1_text, variable=opt1,value=1,text="全選",
                                               command=f1_f1_selectall_event,bg="white",width=11,anchor="w",font=("microsoft yahei",8,"bold"))
    frame1_inside_1_dall_rbtn = tk.Radiobutton(frame1_inside_1_text, variable=opt1,value=2,text="取消全選",
                                               command=f1_f1_dselectall_event,bg="white",width=10,anchor="w",font=("microsoft yahei",8,"bold"))  
    frame1_inside_2_sall_rbtn = tk.Radiobutton(frame1_inside_2_text, var=opt2,value=1,text="全選",
                                               command=f1_f2_selectall_event,bg="white",width=11,anchor="w",font=("microsoft yahei",8,"bold"))
    frame1_inside_2_dall_rbtn = tk.Radiobutton(frame1_inside_2_text, var=opt2,value=2,text="取消全選",
                                               command=f1_f2_dselectall_event,bg="white",width=10,anchor="w",font=("microsoft yahei",8,"bold"))
    frame1_inside_3_sall_rbtn = tk.Radiobutton(frame1_inside_3_text, var=opt3,value=1,text="全選",
                                               command=f1_f3_selectall_event,bg="white",width=11,anchor="w",font=("microsoft yahei",8,"bold"))
    frame1_inside_3_dall_rbtn = tk.Radiobutton(frame1_inside_3_text, var=opt3,value=2,text="取消全選",
                                               command=f1_f3_dselectall_event,bg="white",width=10,anchor="w",font=("microsoft yahei",8,"bold"))
    frame1_inside_1_text.window_create("end", window=frame1_inside_1_sall_rbtn)
    frame1_inside_1_text.window_create("end", window=frame1_inside_1_dall_rbtn)
    frame1_inside_2_text.window_create("end", window=frame1_inside_2_sall_rbtn)
    frame1_inside_2_text.window_create("end", window=frame1_inside_2_dall_rbtn)
    frame1_inside_3_text.window_create("end", window=frame1_inside_3_sall_rbtn)
    frame1_inside_3_text.window_create("end", window=frame1_inside_3_dall_rbtn)
    for i in range(len(frame1_1_check_btn_list)):
        if copy_frame1_1_list[0] in key1_list:
            globals()["f1_in_f1_btn"+str(i)] = tk.Checkbutton(frame1_inside_1_text,text=copy_frame1_1_list[0],var=globals()["f1_in_f1_va"+str(i)],
                                                              width=100,bg="white",fg="blue",anchor="w",onvalue=copy_frame1_1_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame1_inside_1_text.window_create("end", window=globals()["f1_in_f1_btn"+str(i)])
        else:
            globals()["f1_in_f1_btn"+str(i)] = tk.Checkbutton(frame1_inside_1_text,text=copy_frame1_1_list[0],var=globals()["f1_in_f1_va"+str(i)],
                                                              width=100,bg="white",anchor="w",onvalue=copy_frame1_1_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame1_inside_1_text.window_create("end", window=globals()["f1_in_f1_btn"+str(i)])
        if len(copy_frame1_1_list) >1:
            frame1_inside_1_text.insert("end", "\n")
        copy_frame1_1_list.pop(0)
    for i in range(len(frame1_2_check_btn_list)):
        if copy_frame1_2_list[0] in key2_list:
            globals()["f1_in_f2_btn"+str(i)] = tk.Checkbutton(frame1_inside_2_text,text=copy_frame1_2_list[0],var=globals()["f1_in_f2_va"+str(i)],
                                                              width=100,bg="white",fg="blue",anchor="w",onvalue=copy_frame1_2_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame1_inside_2_text.window_create("end", window=globals()["f1_in_f2_btn"+str(i)])
        else:
            globals()["f1_in_f2_btn"+str(i)] = tk.Checkbutton(frame1_inside_2_text,text=copy_frame1_2_list[0],var=globals()["f1_in_f2_va"+str(i)],
                                                              width=100,bg="white",anchor="w",onvalue=copy_frame1_2_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame1_inside_2_text.window_create("end", window=globals()["f1_in_f2_btn"+str(i)])
        if len(copy_frame1_2_list) >1:
            frame1_inside_2_text.insert("end", "\n")
        copy_frame1_2_list.pop(0)
    for i in range(len(frame1_3_check_btn_list)):
        if copy_frame1_3_list[0] in key3_list:
            globals()["f1_in_f3_btn"+str(i)] = tk.Checkbutton(frame1_inside_3_text,text=copy_frame1_3_list[0],var=globals()["f1_in_f3_va"+str(i)],
                                                              width=100,bg="white",fg="blue",anchor="w",onvalue=copy_frame1_3_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame1_inside_3_text.window_create("end", window=globals()["f1_in_f3_btn"+str(i)])
        else:
            globals()["f1_in_f3_btn"+str(i)] = tk.Checkbutton(frame1_inside_3_text,text=copy_frame1_3_list[0],var=globals()["f1_in_f3_va"+str(i)],
                                                              width=100,bg="white",anchor="w",onvalue=copy_frame1_3_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame1_inside_3_text.window_create("end", window=globals()["f1_in_f3_btn"+str(i)])
        if len(copy_frame1_3_list) >1:
            frame1_inside_3_text.insert("end", "\n")
        copy_frame1_3_list.pop(0)
    return frame1_inside_1_sall_rbtn,frame1_inside_2_sall_rbtn,frame1_inside_3_sall_rbtn

# 批量建立資產負債表複選按鈕
def frame2_check_btn():
    key1_1_list = ["現金和短期投資","淨應收款合計","庫存合計","預付費用","其他流動資產合計",
                   "物業/廠房/設備淨總額","商譽淨額","無形資產淨額","長期投資",
                   "長期應收票據","其他長期資產合計","其他資產合計","應付賬款",
                   "應付/應計","應計費用","應付票據/短期債務","長期負債當前應收部分/資本租賃",
                   "其他流動負債合計","長期債務合計","遞延所得稅","少數股東權益","其他負債合計",
                   "可贖回優先股合計","不可贖回優先股淨額","普通股合計","附加資本","保留盈餘(累計虧損)",
                   "普通庫存股","員工持股計劃債務擔保","未實現收益（虧損）","其他權益合計"]
    key1_2_list = ["現金","現金和現金等價物","短期投資","淨交易應收款合計","物業/廠房/設備總額","累計折舊合計","長期債務","資本租賃債務"]
    key2_1_list = ["現金","現金和現金等價物","淨應收款合計","預付費用","物業/廠房/設備淨總額","商譽淨額","無形資產淨額","長期投資",
                   "應收保險","長期應收票據","其他長期資產合計","遞延保單獲得成本","其他資產合計","應付賬款","應付/應計","應計費用",
                   "保單負債","應付票據/短期債務","長期負債當前應收部分/資本租賃","其他流動負債合計","長期債務合計","遞延所得稅",
                   "少數股東權益","其他負債合計","可贖回優先股合計","不可贖回優先股淨額","普通股合計","附加資本","保留盈餘(累計虧損)",
                    "普通庫存股","員工持股計劃債務擔保","未實現收益（虧損）","其他權益合計"]
    key2_2_list = ["物業/廠房/設備總額","累計折舊合計","長期債務","資本租賃債務"]
    key3_1_list = ["銀行應付現金和欠款","其他盈利資產合計","淨貸款","物業/廠房/設備淨總額","商譽淨額","無形資產淨額","長期投資","其他長期資產合計",
                   "其他資產合計","應付賬款","應付/應計","應計費用","存款總額","其他付息負債合計","短期借貸總額","長期負債當前應收部分/資本租賃",
                   "其他流動負債合計","長期債務合計","遞延所得稅","少數股東權益","其他負債合計","可贖回優先股合計","不可贖回優先股淨額","普通股合計",
                   "附加資本","保留盈餘(累計虧損)","普通庫存股","員工持股計劃債務擔保","未實現收益（虧損）","其他權益合計"]
    key3_2_list = ["物業/廠房/設備總額","累計折舊合計","長期債務","資本租賃債務"]
    copy_frame2_1_list = frame2_1_inside_text_list
    copy_frame2_2_list = frame2_2_inside_text_list
    copy_frame2_3_list = frame2_3_inside_text_list
    opt1=tk.IntVar()
    opt2=tk.IntVar()
    opt3 = tk.IntVar()
    frame2_inside_1_sall_rbtn = tk.Radiobutton(frame2_inside_1_text, variable=opt1,value=1,text="全選",
                                               command=f2_f1_selectall_event,bg="white",width=11,anchor="w",font=("microsoft yahei",8,"bold"))
    frame2_inside_1_dall_rbtn = tk.Radiobutton(frame2_inside_1_text, variable=opt1,value=2,text="取消全選",
                                               command=f2_f1_dselectall_event,bg="white",width=10,anchor="w",font=("microsoft yahei",8,"bold"))  
    frame2_inside_2_sall_rbtn = tk.Radiobutton(frame2_inside_2_text, var=opt2,value=1,text="全選",
                                               command=f2_f2_selectall_event,bg="white",width=11,anchor="w",font=("microsoft yahei",8,"bold"))
    frame2_inside_2_dall_rbtn = tk.Radiobutton(frame2_inside_2_text, var=opt2,value=2,text="取消全選",
                                               command=f2_f2_dselectall_event,bg="white",width=10,anchor="w",font=("microsoft yahei",8,"bold"))
    frame2_inside_3_sall_rbtn = tk.Radiobutton(frame2_inside_3_text, var=opt3,value=1,text="全選",
                                               command=f2_f3_selectall_event,bg="white",width=11,anchor="w",font=("microsoft yahei",8,"bold"))
    frame2_inside_3_dall_rbtn = tk.Radiobutton(frame2_inside_3_text, var=opt3,value=2,text="取消全選",
                                               command=f2_f3_dselectall_event,bg="white",width=10,anchor="w",font=("microsoft yahei",8,"bold"))
    frame2_inside_1_text.window_create("end", window=frame2_inside_1_sall_rbtn)
    frame2_inside_1_text.window_create("end", window=frame2_inside_1_dall_rbtn)
    frame2_inside_2_text.window_create("end", window=frame2_inside_2_sall_rbtn)
    frame2_inside_2_text.window_create("end", window=frame2_inside_2_dall_rbtn)
    frame2_inside_3_text.window_create("end", window=frame2_inside_3_sall_rbtn)
    frame2_inside_3_text.window_create("end", window=frame2_inside_3_dall_rbtn)
    for i in range(len(frame2_1_check_btn_list)):
        if copy_frame2_1_list[0] in key1_1_list:
            globals()["f2_in_f1_btn"+str(i)] = tk.Checkbutton(frame2_inside_1_text,text=copy_frame2_1_list[0],var=globals()["f2_in_f1_va"+str(i)],
                                                              width=100,bg="white",fg="blue",anchor="w",onvalue=copy_frame2_1_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame2_inside_1_text.window_create("end", window=globals()["f2_in_f1_btn"+str(i)])
        elif copy_frame2_1_list[0] in key1_2_list:
            globals()["f2_in_f1_btn"+str(i)] = tk.Checkbutton(frame2_inside_1_text,text=copy_frame2_1_list[0],var=globals()["f2_in_f1_va"+str(i)],
                                                              width=100,bg="white",fg="green",anchor="w",onvalue=copy_frame2_1_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame2_inside_1_text.window_create("end", window=globals()["f2_in_f1_btn"+str(i)])
        else:
            globals()["f2_in_f1_btn"+str(i)] = tk.Checkbutton(frame2_inside_1_text,text=copy_frame2_1_list[0],var=globals()["f2_in_f1_va"+str(i)],
                                                              width=100,bg="white",anchor="w",onvalue=copy_frame2_1_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame2_inside_1_text.window_create("end", window=globals()["f2_in_f1_btn"+str(i)])
        if len(copy_frame2_1_list) >1:
            frame2_inside_1_text.insert("end", "\n")
        copy_frame2_1_list.pop(0)
    for i in range(len(frame2_2_check_btn_list)):
        if copy_frame2_2_list[0] in key2_1_list:
            globals()["f2_in_f2_btn"+str(i)] = tk.Checkbutton(frame2_inside_2_text,text=copy_frame2_2_list[0],var=globals()["f2_in_f2_va"+str(i)],
                                                              width=100,bg="white",fg="blue",anchor="w",onvalue=copy_frame2_2_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame2_inside_2_text.window_create("end", window=globals()["f2_in_f2_btn"+str(i)])
        elif copy_frame2_2_list[0] in key2_2_list:
            globals()["f2_in_f2_btn"+str(i)] = tk.Checkbutton(frame2_inside_2_text,text=copy_frame2_2_list[0],var=globals()["f2_in_f2_va"+str(i)],
                                                              width=100,bg="white",fg="green",anchor="w",onvalue=copy_frame2_2_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame2_inside_2_text.window_create("end", window=globals()["f2_in_f2_btn"+str(i)])      
        else:
            globals()["f2_in_f2_btn"+str(i)] = tk.Checkbutton(frame2_inside_2_text,text=copy_frame2_2_list[0],var=globals()["f2_in_f2_va"+str(i)],
                                                              width=100,bg="white",anchor="w",onvalue=copy_frame2_2_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame2_inside_2_text.window_create("end", window=globals()["f2_in_f2_btn"+str(i)])
        if len(copy_frame2_2_list) >1:
            frame2_inside_2_text.insert("end", "\n")
        copy_frame2_2_list.pop(0)
    for i in range(len(frame2_3_check_btn_list)):
        if copy_frame2_3_list[0] in key3_1_list:
            globals()["f2_in_f3_btn"+str(i)] = tk.Checkbutton(frame2_inside_3_text,text=copy_frame2_3_list[0],var=globals()["f2_in_f3_va"+str(i)],
                                                              width=100,bg="white",fg="blue",anchor="w",onvalue=copy_frame2_3_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame2_inside_3_text.window_create("end", window=globals()["f2_in_f3_btn"+str(i)])
        elif copy_frame2_3_list[0] in key3_2_list:
            globals()["f2_in_f3_btn"+str(i)] = tk.Checkbutton(frame2_inside_3_text,text=copy_frame2_3_list[0],var=globals()["f2_in_f3_va"+str(i)],
                                                              width=100,bg="white",fg="green",anchor="w",onvalue=copy_frame2_3_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame2_inside_3_text.window_create("end", window=globals()["f2_in_f3_btn"+str(i)])
        else:
            globals()["f2_in_f3_btn"+str(i)] = tk.Checkbutton(frame2_inside_3_text,text=copy_frame2_3_list[0],var=globals()["f2_in_f3_va"+str(i)],
                                                              width=100,bg="white",anchor="w",onvalue=copy_frame2_3_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame2_inside_3_text.window_create("end", window=globals()["f2_in_f3_btn"+str(i)])
        if len(copy_frame2_3_list) >1:
            frame2_inside_3_text.insert("end", "\n")
        copy_frame2_3_list.pop(0)
    return frame2_inside_1_sall_rbtn,frame2_inside_2_sall_rbtn,frame2_inside_3_sall_rbtn

# 批量建立現金流表複選按鈕
def frame3_check_btn():
    key1_list = [ "折舊/遞耗","攤銷","遞延稅","非現金項目","現金收入","現金支出","現金稅金支出","現金利息支出",
                 "營運資金變動","資本支出","其他投資現金流項目合計","融資現金流項目","發放現金紅利合計",
                 "股票發行（贖回）淨額","債務發行（贖回）淨額"]
    copy_frame3_list = frame3_1_inside_text_list
    opt1=tk.IntVar()
    # opt2=tk.IntVar()
    # opt3=tk.IntVar()
    frame3_inside_1_sall_rbtn = tk.Radiobutton(frame3_inside_1_text, variable=opt1,value=1,text="全選",
                                               command=f3_f1_selectall_event,bg="white",width=11,anchor="w",font=("microsoft yahei",8,"bold"))
    frame3_inside_1_dall_rbtn = tk.Radiobutton(frame3_inside_1_text, variable=opt1,value=2,text="取消全選",
                                               command=f3_f1_dselectall_event,bg="white",width=10,anchor="w",font=("microsoft yahei",8,"bold"))  
    frame3_inside_1_text.window_create("end", window=frame3_inside_1_sall_rbtn)
    frame3_inside_1_text.window_create("end", window=frame3_inside_1_dall_rbtn)
    for i in range(len(frame3_1_check_btn_list)):
        if copy_frame3_list[0] in key1_list:
            globals()["f3_in_f1_btn"+str(i)] = tk.Checkbutton(frame3_inside_1_text,text=copy_frame3_list[0],var=globals()["f3_in_f1_va"+str(i)],
                                                              width=25,bg="white",fg="blue",anchor="w",onvalue=copy_frame3_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame3_inside_1_text.window_create("end", window=globals()["f3_in_f1_btn"+str(i)])

        else:
            globals()["f3_in_f1_btn"+str(i)] = tk.Checkbutton(frame3_inside_1_text,text=copy_frame3_list[0],var=globals()["f3_in_f1_va"+str(i)],
                                                              width=100,bg="white",anchor="w",onvalue=copy_frame3_list[0],offvalue="",font=("microsoft yahei",8,"bold"))
            frame3_inside_1_text.window_create("end", window=globals()["f3_in_f1_btn"+str(i)])
        if len(copy_frame3_list) >1:
            frame3_inside_1_text.insert("end", "\n")
        copy_frame3_list.pop(0)
    return frame3_inside_1_sall_rbtn

# 自適應視窗中心
align_center(window,1000,700)
window.update_idletasks() 

"""instance compoments"""
country_lb = tk.Label(window,text="股票代號/國家： ",font=("新細明體",12),anchor="e")
country_lb.place(relx=0.01,rely=0.05,relwidth=0.15)

country_btn = ttk.Combobox(window,value=country_list)
country_btn.place(relx=0.16,rely=0.05,relwidth=0.3)
country_btn.current(0)

market_lb = tk.Label(window,text="市場交易所： ",font=("新細明體",12),anchor="e")
market_lb.place(relx=0.01,rely=0.12,relwidth=0.15)

market_btn = ttk.Combobox(window,state="readonly")
market_btn.place(relx=0.16,rely=0.12,relwidth=0.3)

category_lb = tk.Label(window,text="股票類別： ",font=("新細明體",12),anchor="e")
category_lb.place(relx=0.01,rely=0.19,relwidth=0.15)

category_btn = ttk.Combobox(window,state= "readonly")
category_btn.place(relx=0.16,rely=0.19,relwidth=0.3)

start_date_lb = tk.Label(window,text="開始日期： ",font=("新細明體",12),anchor="e")
start_date_lb.place(relx=0.01,rely=0.26,relwidth=0.15)

year,month,day = get_dateconfig()
start_date_btn = DateEntry(window,width=10,background='darkblue', year=year,month=month,day=day,
                           foreground="white", borderwidth=2, locale = "en_us", date_pattern ="yyyy/mm/dd")
start_date_btn.place(relx=0.16,rely=0.26,relwidth=0.3)

end_date_lb = tk.Label(window,text="結束日期： ",font=("新細明體",12),anchor="e")
end_date_lb.place(relx=0.01,rely=0.33,relwidth=0.15)

end_date_btn = DateEntry(window,width=10,background='darkblue', foreground="white", borderwidth=2, locale = "en_us", date_pattern ="yyyy/mm/dd")
end_date_btn.place(relx=0.16,rely=0.33,relwidth=0.3)

stext = ScrolledText(window,bg="white",selectbackground="blue")
stext.config(state="disabled",font=("新細明體",13))
stext.place(relx=0.01,rely=0.4,relwidth=0.49,relheight=0.2)

window_spea = ttk.Separator(window,orient="horizontal")
window_spea.place(relx=0,rely=0.62,relwidth=1)                                  

samplepath_lb = tk.Label(window,text="合併範例檔案： ",font=("新細明體",12),anchor="e")
samplepath_lb.place(relx=0.01,rely=0.635,relwidth=0.15)

samplepath_btn = ttk.Button(window,text="選擇檔案",command=samplepath_event)
samplepath_btn.place(relx=0.16,rely=0.63,relheight=0.05)

samplepath_text = tk.Text(window)
samplepath_text.config(state="disabled")
samplepath_text.place(relx=0.03,rely=0.7,relwidth=0.4,relheight=0.04)

downloadpath_lb = tk.Label(window,text="個股下載路徑： ",font=("新細明體",12),anchor="e")
downloadpath_lb.place(relx=0.01,rely=0.755,relwidth=0.15)

downloadpath_btn = ttk.Button(window,text="選擇路徑",command=downloadpath_event)
downloadpath_btn.place(relx=0.16,rely=0.75)

downloadpath_text = tk.Text(window)
downloadpath_text.config(state="normal")
downloadpath_text.insert(1.0,"C:/")
downloadpath_text.config(state="disabled")
downloadpath_text.place(relx=0.03,rely=0.82,relwidth=0.4,relheight=0.04)

bulkdownloadpath_lb = tk.Label(window,text="批量下載路徑： ",font=("新細明體",12),anchor="e")
bulkdownloadpath_lb.place(relx=0.01,rely=0.875,relwidth=0.15)

bulkdownloadpath_btn = ttk.Button(window,text="選擇路徑",command=downloadpath2_event)
bulkdownloadpath_btn.place(relx=0.16,rely=0.87)

bulkdownloadpath_text = tk.Text(window)
bulkdownloadpath_text.config(state="normal")
bulkdownloadpath_text.insert(1.0,"C:/")
bulkdownloadpath_text.config(state="disabled")
bulkdownloadpath_text.place(relx=0.03,rely=0.94,relwidth=0.4,relheight=0.04)

stocklistpath_lb = tk.Label(window,text="指定清單檔案： ",font=("新細明體",12),anchor="e")
stocklistpath_lb.place(relx=0.51,rely=0.635,relwidth=0.15)

stocklistpath_btn = ttk.Button(window,text="選擇檔案",command=stocklistpath_event)
stocklistpath_btn.place(relx=0.66,rely=0.63)

stocklistpath_text = tk.Text(window)
stocklistpath_text.config(state="disabled")
stocklistpath_text.place(relx=0.53,rely=0.7,relwidth=0.4,relheight=0.04)

crawlerstocklist_btn = ttk.Button(window,text="下載個股清單",command=crawlerstocklist_btn_event)
crawlerstocklist_btn.place(relx=0.53,rely=0.8)

stocktimer_lb = tk.Label(window,text="",font=("新細明體",10),anchor="e")
stocktimer_lb.place(relx=0.53,rely=0.88)

bulkcrawler_btn = ttk.Button(window,text="下載清單資料",command=bulkcrawler_btn_event)
bulkcrawler_btn.place(relx=0.68,rely=0.8)

bulk_start_btn = ttk.Button(window,text="暫停下載",command=stopcrawler_btn_event)
bulk_start_btn.place(relx=0.68,rely=0.87)

singlecrawler_btn = ttk.Button(window,text="下載目標個股",command=singlecrawler_btn_event)
singlecrawler_btn.place(relx=0.83,rely=0.8)

progress = ttk.Progressbar(window,style="red.Horizontal.TProgressbar",orient="horizontal",mode="determinate")
progress.place(relx=0.53,rely=0.94,relwidth=0.43)

# window notebook&frame
notebook = ttk.Notebook(window)
notebook.place(relx=0.5,rely=0,relwidth=0.5,relheight=0.6)          
frame_1 = tk.Frame(notebook)
frame_2 = tk.Frame(notebook)
frame_3 = tk.Frame(notebook)

# frame1~4 grid
notebook.add(frame_1,text="損益表",padding=10)
notebook.add(frame_2,text="資產負債表",padding=10)
notebook.add(frame_3,text="現金流",padding=10)

# frame1 configure
frame_1_notebook = ttk.Notebook(frame_1)                                        
frame_1_notebook.place(relx=0,rely=0,relwidth=1,relheight=1)                   
frame1_inside_1 = tk.Frame(frame_1_notebook)                                    
frame1_inside_1.place(relx=0,rely=0,relwidth=1,relheight=1)                    
frame1_inside_1_scollbar = tk.Scrollbar(frame1_inside_1)                       
frame1_inside_1_scollbar.place(relx=0.95,relwidth=0.05,relheight=1)             
frame1_inside_1_text = tk.Text(frame1_inside_1)                                
frame1_inside_1_text.place(relx=0,rely=0,relwidth=0.95,relheight=1)             
frame1_inside_1_text.configure(state="disabled")
frame1_inside_1_scollbar.config(command=frame1_inside_1_text.yview)
frame1_inside_1_text.config(yscrollcommand=frame1_inside_1_scollbar.set)
frame1_inside_2 = tk.Frame(frame_1_notebook)
frame1_inside_2.place(relx=0,rely=0,relwidth=1,relheight=1)
frame1_inside_2_scollbar = tk.Scrollbar(frame1_inside_2)                        
frame1_inside_2_scollbar.place(relx=0.95,relwidth=0.05,relheight=1)             
frame1_inside_2_text = tk.Text(frame1_inside_2)                                 
frame1_inside_2_text.place(relx=0,rely=0,relwidth=0.95,relheight=1)             
frame1_inside_2_text.configure(state="disabled")
frame1_inside_2_scollbar.config(command=frame1_inside_2_text.yview)
frame1_inside_2_text.config(yscrollcommand=frame1_inside_2_scollbar.set)
frame1_inside_3 = tk.Frame(frame_1_notebook)
frame1_inside_3.place(relx=0,rely=0,relwidth=1,relheight=1)
frame1_inside_3_scollbar = tk.Scrollbar(frame1_inside_3)                        
frame1_inside_3_scollbar.place(relx=0.95,relwidth=0.05,relheight=1)             
frame1_inside_3_text = tk.Text(frame1_inside_3)                                 
frame1_inside_3_text.place(relx=0,rely=0,relwidth=0.95,relheight=1)             
frame1_inside_3_text.configure(state="disabled")
frame1_inside_3_scollbar.config(command=frame1_inside_3_text.yview)
frame1_inside_3_text.config(yscrollcommand=frame1_inside_3_scollbar.set)
f1_var_name()
frame1_inside_1_sall_rbtn,frame1_inside_2_sall_rbtn,frame1_inside_3_sall_rbtn = frame1_check_btn()
frame_1_notebook.add(frame1_inside_1,text="一般類別")                              
frame_1_notebook.add(frame1_inside_2,text="保險類別")     
frame_1_notebook.add(frame1_inside_3,text="銀行類別")                    

# frame2 configure
frame_2_notebook = ttk.Notebook(frame_2)
frame_2_notebook.place(relx=0,rely=0,relwidth=1,relheight=1)       
frame2_inside_1 = tk.Frame(frame_2_notebook)
frame2_inside_1.place(relx=0,rely=0,relwidth=1,relheight=1) 
frame2_inside_1_scollbar = tk.Scrollbar(frame2_inside_1)                        
frame2_inside_1_scollbar.place(relx=0.95,relwidth=0.05,relheight=1)             
frame2_inside_1_text = tk.Text(frame2_inside_1)                                 
frame2_inside_1_text.place(relx=0,rely=0,relwidth=0.95,relheight=1)             
frame2_inside_1_text.configure(state="disabled")
frame2_inside_1_scollbar.config(command=frame2_inside_1_text.yview)
frame2_inside_1_text.config(yscrollcommand=frame2_inside_1_scollbar.set)
frame2_inside_2 = tk.Frame(frame_2_notebook)
frame2_inside_2.place(relx=0,rely=0,relwidth=1,relheight=1)
frame2_inside_2_scollbar = tk.Scrollbar(frame2_inside_2)                        
frame2_inside_2_scollbar.place(relx=0.95,relwidth=0.05,relheight=1)             
frame2_inside_2_text = tk.Text(frame2_inside_2)                                 
frame2_inside_2_text.place(relx=0,rely=0,relwidth=0.95,relheight=1)             
frame2_inside_2_text.configure(state="disabled")
frame2_inside_2_scollbar.config(command=frame2_inside_2_text.yview)
frame2_inside_2_text.config(yscrollcommand=frame2_inside_2_scollbar.set)
frame2_inside_3 = tk.Frame(frame_2_notebook)
frame2_inside_3.place(relx=0,rely=0,relwidth=1,relheight=1)
frame2_inside_3_scollbar = tk.Scrollbar(frame2_inside_3)                        
frame2_inside_3_scollbar.place(relx=0.95,relwidth=0.05,relheight=1)             
frame2_inside_3_text = tk.Text(frame2_inside_3)                                 
frame2_inside_3_text.place(relx=0,rely=0,relwidth=0.95,relheight=1)             
frame2_inside_3_text.configure(state="disabled")
frame2_inside_3_scollbar.config(command=frame2_inside_3_text.yview)
frame2_inside_3_text.config(yscrollcommand=frame2_inside_3_scollbar.set)
f2_var_name()
frame2_inside_1_sall_rbtn,frame2_inside_2_sall_rbtn,frame2_inside_3_sall_rbtn = frame2_check_btn()
frame_2_notebook.add(frame2_inside_1,text="一般類別")
frame_2_notebook.add(frame2_inside_2,text="保險類別")
frame_2_notebook.add(frame2_inside_3,text="銀行類別")  

# frame3 configure
frame_3_notebook = ttk.Notebook(frame_3)
frame_3_notebook.place(relx=0,rely=0,relwidth=1,relheight=1)   
frame3_inside_1 = tk.Frame(frame_3_notebook)
frame3_inside_1.place(relx=0,rely=0,relwidth=1,relheight=1) 
frame3_inside_1_scollbar = tk.Scrollbar(frame3_inside_1)                        
frame3_inside_1_scollbar.place(relx=0.95,relwidth=0.05,relheight=1)             
frame3_inside_1_text = tk.Text(frame3_inside_1)                                 
frame3_inside_1_text.place(relx=0,rely=0,relwidth=0.95,relheight=1)             
frame3_inside_1_text.configure(state="disabled")
frame3_inside_1_scollbar.config(command=frame3_inside_1_text.yview)
frame3_inside_1_text.config(yscrollcommand=frame3_inside_1_scollbar.set)     
f3_var_name()
frame3_inside_1_sall_rbtn = frame3_check_btn()
frame_3_notebook.add(frame3_inside_1,text="一般類別")

"""event bind configure"""
start_date_btn.bind("<<DateEntrySelected>>", start_date_btn_event)
# end_date_btn.bind("<<DateEntrySelected>>", end_date_btn_event)
country_btn.bind("<<ComboboxSelected>>", country_btn_event)
country_btn.bind("<Return>",country_btn_enter_event)

"""default checkbox select"""
f1_1config,f1_2config,f1_3config,f2_1config,f2_2config,f2_3config,f3config = get_checkboxconfig()
for i in range(len(frame1_1_check_btn_list)):
    if f1_1config[i]:
        globals()["f1_in_f1_btn"+str(i)].select()
    else:
        globals()["f1_in_f1_btn"+str(i)].deselect()
for i in range(len(frame1_2_check_btn_list)):
    if f1_2config[i]:
        globals()["f1_in_f2_btn"+str(i)].select()
    else:
        globals()["f1_in_f2_btn"+str(i)].deselect()
for i in range(len(frame1_3_check_btn_list)):
    if f1_3config[i]:
        globals()["f1_in_f3_btn"+str(i)].select()
    else:
        globals()["f1_in_f3_btn"+str(i)].deselect()
for i in range(len(frame2_1_check_btn_list)):
    if f2_1config[i]:
        globals()["f2_in_f1_btn"+str(i)].select()
    else:
        globals()["f2_in_f1_btn"+str(i)].deselect()
for i in range(len(frame2_2_check_btn_list)):
    if f2_2config[i]:
        globals()["f2_in_f2_btn"+str(i)].select()
    else:
        globals()["f2_in_f2_btn"+str(i)].deselect()
for i in range(len(frame2_3_check_btn_list)):
    if f2_3config[i]:
        globals()["f2_in_f3_btn"+str(i)].select()
    else:
        globals()["f2_in_f3_btn"+str(i)].deselect()
for i in range(len(frame3_1_check_btn_list)):
    if f3config[i]:
        globals()["f3_in_f1_btn"+str(i)].select()
    else:
        globals()["f3_in_f1_btn"+str(i)].deselect()        

"""default radiobox select"""
frame1_inside_1_sall_rbtn.select()
frame1_inside_2_sall_rbtn.select()
frame1_inside_3_sall_rbtn.select()
frame2_inside_1_sall_rbtn.select()
frame2_inside_2_sall_rbtn.select()
frame2_inside_3_sall_rbtn.select()
frame3_inside_1_sall_rbtn.select()

"""get path temp"""
single,bulk = get_pathconfig()
downloadpath_text.config(state="normal")
downloadpath_text.delete(1.0,"end")
downloadpath_text.insert(1.0, single)
downloadpath_text.config(state="disable")
bulkdownloadpath_text.config(state="normal")
bulkdownloadpath_text.delete(1.0,"end")
bulkdownloadpath_text.insert(1.0, bulk)
bulkdownloadpath_text.config(state="disable")

# window.update()
# wd,hi = notebook.winfo_width(), notebook.winfo_height()
# print(wd,hi)

if __name__ == '__main__':
    window.mainloop()
