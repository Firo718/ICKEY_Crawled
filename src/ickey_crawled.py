#!/usr/bin/env python
# -*- coding: utf-8 -*-

from selenium import webdriver
from bs4 import BeautifulSoup
import csv
import xlrd
import types

def IC(ic_name):
    driver = webdriver.PhantomJS()
    url = 'http://search.ickey.cn/?keyword=' + ic_name +'&num='
    print(url)
    driver.get(url)
    soup = BeautifulSoup(driver.page_source,"html.parser")
    driver.quit()
    x = soup.body.find_all("div", "search-data-item")
    print("*********************************************")
    x1 = []
    x2 = []
    for i in range(len(x)):
        x2 = company(x[i],ic_name)
        x1 = x1 + x2
    print(x1)
    return x1

def company(search_data_item,ic_name):
    company_name = search_data_item['data-domid'][4:]
    print(company_name)
    table_tbody_data_rowid = search_data_item.table.tbody.find_all("tr", attrs={"data-rowid": True})
    x1 = []
    x2 = []
    for i in range(len(table_tbody_data_rowid)):
        x2 = final_level(table_tbody_data_rowid[i].find_all("td"),company_name,ic_name)
        x1 = x1 + x2
    print(x1)
    return x1
    print("*********************************************")

def final_level(tbody_tr_datarowid_td,company_name,ic_name):
    a = stock_value(tbody_tr_datarowid_td)
    b = MOQ_value(tbody_tr_datarowid_td)
    c = RMB_value(tbody_tr_datarowid_td)
    print("库存：", a)
    print("起订量：", b)
    print("价格：", c)
    x = []
    for i in range(len(b)):
        x.append([ic_name,company_name,a,b[i],c[i]])
    return x

def stock_value(tbody_tr_datarowid_td):
    stock = tbody_tr_datarowid_td[4].font.string
    return stock

def MOQ_value(tbody_tr_datarowid_td):
    MOQ = tbody_tr_datarowid_td[5].find_all("div",attrs={"data-value":True})
    numofMOQ = []
    for i in range(len(MOQ)):
        numofMOQ.append(MOQ[i].string)
    return numofMOQ

def RMB_value(tbody_tr_datarowid_td):
    RMB = tbody_tr_datarowid_td[7].div.find_all("div")
    priceofRMB = []
    for i in range(len(RMB)):
        priceofRMB.append(RMB[i].string[1:])
    return priceofRMB


#data = IC('AD9361BBCZ') + IC('XC7A100T-2FGG676I')

book = xlrd.open_workbook('x.xls')

print('sheet num:', book.nsheets)
print('sheet name', book.sheet_names())
sheet1 = book.sheet_by_index(0)
col_data = sheet1.col_values(0)[1:]
print(col_data)

data = []
for i in range(len(col_data)):
    data = data + IC(col_data[i])

with open('test.csv', 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['IC_name','Company','stock','MOQ','price:人民币含税'])
    writer.writerows(data)
    csvfile.close()