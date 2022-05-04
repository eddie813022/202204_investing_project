from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
import pandas as pd
import requests
import time

user_agent = UserAgent()
headers = {'User-Agent':user_agent.random}
url = "https://hk.investing.com/"
res = requests.get(url, headers=headers)
res.encoding = "UTF-8"
country_xml = BeautifulSoup(res.text,"lxml").find(id="countriesUL").find_all("li")
country_list = [ (i.text).replace("\n","") for i in country_xml ]

options = webdriver.FirefoxOptions()
options.add_argument("--disable-notifications")
options.add_argument("windows-size=1920,1080")
# options.add_argument("--headless")
service = Service("geckodriver.exe")
driver = webdriver.Firefox(service=service,options=options)
driver.maximize_window()
driver.get("https://hk.investing.com/")
WebDriverWait(driver,3).until(EC.presence_of_element_located((By.CSS_SELECTOR,".selectWrap.js-filter-dropdown[data-filter-type='country']")))
scoll_posision = driver.find_element(By.CSS_SELECTOR,".homepageWidget.stockScreenerWrap")
country_btn = driver.find_element(By.CSS_SELECTOR,".selectWrap.js-filter-dropdown[data-filter-type='country']").find_element(By.CSS_SELECTOR,".newBtnDropdown.noHover")
country_input = driver.find_element(By.CSS_SELECTOR,".selectWrap.js-filter-dropdown[data-filter-type='country']").find_element(By.CSS_SELECTOR,".js-search-input.inputDropDown")
driver.execute_script("arguments[0].scrollIntoView()",scoll_posision) 

try:
    country_btn.click()
    country_input.send_keys("台灣")
    time.sleep(1)
    country_input.send_keys(Keys.DOWN)
    country_input.send_keys(Keys.ENTER)
    WebDriverWait(driver,3).until(EC.presence_of_element_located((By.ID,"exchangesUL")))
    # print(tu.get_attribute("innerHTML"))
except Exception as e :
    print(e)
# country = tu.find_element(By.CSS_SELECTOR,"li[data-value='24']").click()