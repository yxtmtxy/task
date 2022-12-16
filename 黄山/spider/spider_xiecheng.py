import requests
import json
import re
import pandas as pd
import time
from openpyxl import load_workbook

workbook = load_workbook(filename ='xiecheng1.xlsx')
sheet = workbook.active

url='https://m.ctrip.com/restapi/soa2/13444/json/getCommentCollapseList'
response = requests.get(url)
pagetext = response.text
a=300
i=2992
while(a<350):
   data = {
     "arg": {
       "channelTyp": "2",
       "collapseType":"0",
       "ommentTagId": "0",
       "pageIndex": a,
       "pageSize": "10",
       "poiId": "87869",
       "sourceType":"1",
       "sortType":"3",
       "starType":"0"
     }

   }
   headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36',
    'cookie':'GUID=09031053212378896635; nfes_isSupportWebP=1; _bfaStatusPVSend=1; MKT_CKID=1662967369190.nrycs.gcl7; MKT_CKID_LMT=1662967369190; _RF1=60.168.252.161; _RSG=_4qy6UNcbh9f8g6xPUQI6B; _RDG=28f5a52ca272362b1a21ee895162fd6c0e; _RGUID=9ae59977-86f3-413f-b192-69e67b1aa229; _bfa=1.1662967366090.3h37i6.1.1662967366090.1662970044295.2.2.1; _bfs=1.1; _ubtstatus=%7B%22vid%22%3A%221662967366090.3h37i6%22%2C%22sid%22%3A2%2C%22pvid%22%3A2%2C%22pid%22%3A0%7D; _bfi=p1%3D290510%26p2%3D290510%26v1%3D2%26v2%3D1; _bfaStatus=success; _jzqco=%7C%7C%7C%7C1662967369606%7C1.1542856578.1662967369189.1662967369189.1662970048129.1662967369189.1662970048129.0.0.0.2.2; __zpspc=9.2.1662970048.1662970048.1%234%7C%7C%7C%7C%7C%23'
   }
   r = requests.post(url=url, data=json.dumps(data), headers=headers)
   j = r.json()#转为json
   c = j['result']
   num = [i for i in range(len(c['items']))]
   print(a)
   for x in num:
           scores=c['items'][x]['scores']
           i += 1
           str1 = 'A' + str(i)
           str2 = 'B' + str(i)
           str3 = 'C' + str(i)
           str4 = 'D' + str(i)


           print(c['items'][x]['content'])
           sheet[str4].value = c['items'][x]['content']

           workbook.save(filename='xiecheng1.xlsx')
   time.sleep(1)
   a += 1




