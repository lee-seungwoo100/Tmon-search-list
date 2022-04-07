import time
import re
from datetime import datetime
from typing import ItemsView
from selenium import webdriver
import requests
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import openpyxl

search_product ='고구마'


excel_file = openpyxl.Workbook()
excel_sheet2 = excel_file.active
excel_sheet2.append(['순위','상품명', '가격', '리뷰수','판매수','링크'])
excel_sheet2.title = '티몬'
excel_sheet2.column_dimensions['B'].width = 80
excel_sheet2.column_dimensions['C'].width = 15
excel_sheet2.column_dimensions['D'].width = 15
excel_sheet2.column_dimensions['E'].width = 15


options = webdriver.ChromeOptions()
#options.add_argument('headless')
options.add_argument("disable-gpu")
options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.84 Safari/537.36')

browser = webdriver.Chrome(options=options) #'./chromdriver.exe'


# 티몬 특정키워드로 셀레니움 정보창 정보 가져오기


browser.get('https://www.tmon.co.kr/')
browser.implicitly_wait(10)
browser.find_element_by_name('keyword').send_keys(search_product)
browser.find_element_by_class_name('btn_search').click()
time.sleep(3)

items = browser.find_elements(By.CSS_SELECTOR,'div.deallist_wrap > ul > li')
for index, item in enumerate (items,start=1):
    productname = item.find_element(By.CSS_SELECTOR,'div.deal_info > p').text
    productprice = item.find_element(By.CSS_SELECTOR,'div.deal_info span.price i.num').text
    try:
        productreview = item.find_element(By.CSS_SELECTOR,'div.deal_info span.grade_average_count > span.num').text
    except:
        productreview = '리뷰없음'
    try:
        productbuy = item.find_element(By.CSS_SELECTOR,'div.deal_info span.buy_count').text[:-2]
    except:
        productbuy = '구매없음'
    productlink = item.find_element(By.CSS_SELECTOR,'li.item > a').get_attribute('href')

    print(productname,productprice,productreview,productbuy,productlink)
    excel_sheet2.append([index,productname,productprice,productreview,productbuy,productlink])
    excel_sheet2.cell(row=index, column=6).hyperlink = productlink

cell_A1 = excel_sheet2['A1'] # 셀 선택하기
cell_A1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
cell_A1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
# 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

cell_B1 = excel_sheet2['B1'] # 셀 선택하기
cell_B1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
cell_B1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
# 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

cell_C1 = excel_sheet2['C1'] # 셀 선택하기
cell_C1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
cell_C1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
# 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

cell_D1 = excel_sheet2['D1'] # 셀 선택하기
cell_D1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
cell_D1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
# 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

cell_E1 = excel_sheet2['E1'] # 셀 선택하기
cell_E1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
cell_E1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
# 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

cell_F1 = excel_sheet2['F1'] # 셀 선택하기
cell_F1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
cell_F1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
# 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors            


excel_file.save(search_product+'베스트'+'.xlsx')
excel_file.close()
browser.quit()
