import requests
import json
import re
import pandas as pd
import time
from openpyxl import load_workbook

workbook = load_workbook(filename ='meituan6.xlsx')
sheet = workbook.active

url='https://www.meituan.com/ptapi/poi/getcomment'
response = requests.get(url)
pagetext = response.text
a=13000
while a<16000:
      print(a)
      param = {
               'id': '2258206',
               'offset': a,   #从第几个评论爬取
               'pageSize': '10',#一次爬取次数
               'mode': '0',
               'starRange':'',
               'userId':'',
               'sortType': '1',
               }
      headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36',
                 'Cookie': '_lxsdk_cuid = 182f8cc8744c8 - 07a29be1be9e93 - 26021d51 - 127184 - 182f8cc8744c8;uuid = 076cb82b076b4a14b380.1667381254.1.0.0'
                 }
      r = requests.get(url=url, params=param, headers=headers)
      j = r.json()
      num = [i for i in range(len(j['comments']))]
      i=a
      for x in num:
          i +=1
          print(j['comments'][x]['star'])
          print(j['comments'][x]['comment'])
          str1 = 'A' + str(i)
          str2 = 'B' + str(i)
          sheet[str1].value = j['comments'][x]['star']
          sheet[str2].value = j['comments'][x]['comment']
          workbook.save(filename ='meituan6.xlsx')
      time.sleep(1)
      a+=10
