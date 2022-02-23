#excel.py
from itertools import count
from lib2to3 import refactor
from openpyxl import Workbook

#load
import pandas as pd
def findFile(filename):
    df = pd.read_excel(filename)
    return df['keyword'].tolist()

#search
import sys
import io
import time

sys.stdout = io.TextIOWrapper(sys.stdout.detach(), encoding = 'utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.detach(), encoding = 'utf-8')


#검색하기
import urllib.parse
import urllib.request
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from bs4 import BeautifulSoup
def search(word,findWord,totalCnt):
    i=1
    rank=1
    longStr=""
    while(True):
        baseUrl="https://search.shopping.naver.com/search/all?frm=NVSHATC&pagingIndex="+str(i)+"&pagingSize=40&productSet=total&query="
        url=baseUrl+urllib.parse.quote_plus(word)
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('headless')
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        
        driver.get(url)
        driver.implicitly_wait(10)
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
        driver.implicitly_wait(10)
        time.sleep(1)
        soup=BeautifulSoup(driver.page_source,'html.parser')
        driver.quit()
        li_item=soup.select(".list_basis>div>div")
        theEnd=False
        for item in li_item:
            aTag=item.select_one('.basicList_mall_title__3MWFY a')
            if(len(item.select('.ad_ad_stk__12U34'))>=1) :
                continue
            elif(aTag.text.find("쇼핑몰별 최저가")>=0):
                continue
            elif(aTag.text.find("브랜드 카탈로그")>=0):
                continue
            else :
                if(aTag.text.find(findWord)>=0):
                    theEnd=True
                    break
                longStr=longStr+"==="+aTag.text
                rank=rank+1   
        if(theEnd):
            break
        i=i+1
        if(rank>totalCnt):
            rank="{}순위에 벗어남".format(totalCnt)
            break        
    print("'{}'로 검색하여 '{}'이 단어의 순위는 '{}'입니다.".format(word,findWord,rank))
#search(검색어,찾고자하는 단어,순위)
search('사과','한나옹산',50);