import requests
import json
import re
import pandas as pd
import time
from openpyxl import load_workbook

workbook = load_workbook(filename ='sr.xlsx')
sheet = workbook.active

url=''
response = requests.get(url)
pagetext = response.text
a=0
while a<3110:
      param = {

               }
      headers = {
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
          workbook.save(filename ='meituan1.xlsx')
      time.sleep(1)
      a+=10
