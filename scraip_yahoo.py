import re
import sys
import time
import urllib

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait

import main_form
from search_item import SearchItem
import mylogger

# ログの定義
logger = mylogger.setup_logger(__name__)

def create_driver():
    # chromeドライバーのパス
    chrome_path = "./driver/chromedriver.exe"

    # Selenium用オプション
    op = Options()
    op.add_argument('--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/12.0 Safari/605.1.15')
    op.add_experimental_option("excludeSwitches", ["enable-automation"])
    op.add_experimental_option('useAutomationExtension', False)

    #op.add_argument("--disable-gpu");
    #op.add_argument("--disable-extensions");
    #op.add_argument("--proxy-server='direct://'");
    #op.add_argument("--proxy-bypass-list=*");
    #op.add_argument("--start-maximized");
    #op.add_argument("--headless");
    #driver = webdriver.Chrome(chrome_options=op)
    driver = webdriver.Chrome(executable_path=chrome_path, chrome_options=op)
    return driver


def get_chiebukuro_url(driver, url, item: SearchItem):

    # 人間に近くするために、Wait
    main_form.wait_randam_sec(5.0, 8.0)
    driver.get(url)

    # 検索
    input_sel = 'input.SearchBox__searchInput'
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, input_sel))
    )
    # 人間に近くするために、Wait
    main_form.wait_randam_sec(1.2, 2.5)

    driver.find_elements_by_css_selector(input_sel)[0].send_keys(item.search_keyword)

    # 人間に近くするために、Wait
    main_form.wait_randam_sec(0.5, 3.0)
    driver.find_elements_by_css_selector(input_sel)[0].send_keys(Keys.ENTER)

    WebDriverWait(driver, 30).until(
        EC.presence_of_all_elements_located
    )
    div_sel = "div.AnswerChiebukuro"
    soup = BeautifulSoup(driver.page_source, features="html.parser")
    chiebukuro_ele = soup.select(div_sel)
    if not chiebukuro_ele:
        item.search_url_no1 = '枠無し'
        item.search_url_no2 = '枠無し'
        item.search_url_no3 = '枠無し'
        return item

    a_list = chiebukuro_ele[0].select("a")
    if len(a_list) >= 2:
        item.search_url_no1 = a_list[1].get("href")
    if len(a_list) >= 3:
        item.search_url_no2 = a_list[2].get("href")
    if len(a_list) >= 4:
        item.search_url_no3 = a_list[3].get("href")

    return item
