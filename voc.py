# -*- coding: utf-8 -*-
"""
Created on Fri Dec 15 21:16:42 2017

@author: Hu_Zi
"""

from bs4 import BeautifulSoup
import urllib
#import openpyxl
from openpyxl import load_workbook
name = input("input file name here")
def finddef(voc):
    try:
        source=urllib.request.urlopen("https://www.collinsdictionary.com/dictionary/english/"+voc)
    #print(source.read())
        soup=BeautifulSoup(source.read(),"lxml")
        define=soup.find('div','def')
        explain=(define.get_text())
        return explain
    except:
        return "error"
#    print(explain)
#voc=input()
#finddef(voc)
table=load_workbook('F:\\单词默写\\'+name+'\\'+name+'.xlsx')#file name
sheet1=table['Sheet1']
sheet2=table['Sheet2']
print('start working',end='')
for i in range(1,91):
    voc=sheet1['A'+str(i)].value
#    print(voc)
    sheet2['A'+str(i)].value=voc
    sheet2['B'+str(i)].value=finddef(voc)
    print('#',end='')
#    print(sheet2['B'+str(i)].value)
table.save('F:\\单词默写\\'+name+'\\'+name+'.xlsx')#file name
print('\nfinshed.')
#print(table.get_sheet_names())