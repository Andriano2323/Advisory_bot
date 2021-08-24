# -*- coding: utf-8 -*-
"""
Created on Sun Aug 22 20:29:26 2021

@author: andre
"""

import bs4 
import requests
from bs4 import BeautifulSoup
def parsePrice():
    r = requests.get('https://finance.yahoo.com/quote/FB?p=FB')
    soup = bs4.BeautifulSoup(r.text,"html.parser")    
    price = soup.find_all('div',{'class':'My(6px) Pos(r) smartphone_Mt(6px)'})[0].find('span').text
    return price


