
# -*- coding:utf-8 -*-
from http.client import responses
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import Cell
from openpyxl.styles import Font
from time import sleep, perf_counter, perf_counter
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
from pathlib import Path
from datetime import date, datetime
from lxml.etree import ParserError
import pandas as pd
import requests
import random

user_agent = UserAgent()
headers = {'User-Agent': user_agent.random}



"""股息URL測試"""

def get_dividend(url):
    if "?cid=" in url:
        f_index = url.find("?cid=")
        l_index = url[f_index:]
        url = url[:f_index]+"-dividends"+l_index
        res = requests.get(url, headers=headers)
        res.encoding = "UTF-8"
        xml = BeautifulSoup(res.text, "lxml")
    else:
        url = url+"-dividends"
        res = requests.get(url, headers=headers)
        res.encoding = "UTF-8"
        xml = BeautifulSoup(res.text, "lxml")
    finaldf = pd.DataFrame()
    df = pd.read_html(str(xml))
    for i in df:
        if "除息日" in i.columns:
            # if "除息日" in str(i.columns[0]):
            finaldf = i
    frame_to_rows = list(dataframe_to_rows(finaldf, index=False, header=False))
    print(frame_to_rows)

def get_more_dividend(pair_id):
    url = "https://hk.investing.com/equities/MoreDividendsHistory"
    stock_data = { "pairID": pair_id, "last_timestamp": 1309910400 }
    headers = { "User-Agent": user_agent.random,
                "x-requested-with": "XMLHttpRequest" }
    res = requests.post(url, headers=headers, data=stock_data)
    res.encoding = "UTF-8"
    history_row = res.json()["historyRows"]
    xml = BeautifulSoup(history_row, "html.parser")
    print(xml)
    # date_ = xml.find_all("td", class_="left first")
    # date = [ i.text for i in date_ ]
    # print(date)

# 1498608000
# 1309910400 = 188,697,600
# 1121731200 = 188,179,200
# 1000000000

# 1502409600
# 1313107200



# url = "https://hk.investing.com/equities/formosa-petro"
# response =get_dividend(url=url)
response = get_more_dividend(pair_id="103588")
# 103731