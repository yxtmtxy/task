import requests
import json
import re
import pandas as pd
import time
from openpyxl import load_workbook

workbook = load_workbook(filename ='qunaer1.xlsx')
sheet = workbook.active

url='http://piao.qunar.com/ticket/detailLight/sightCommentList.json'
response = requests.get(url)
pagetext = response.text

a=1
i=379
while(a<23):
     param = {
               'sightId': '16176',
               'index': a,
               'page': a,
               'pageSize': '10',
               'tagType': '3'
     }
     headers = {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36',
          'cookie': 'SECKEY_ABVK=oM6TZxG/R/b9aK3c9yTN23L4/36haWPaubpGskpoY1Y%3D; BMAP_SECKEY=q_7oQU1tMlv2SxxraeoRKj9PnnZ_Npb4bi73tiHHscRjXfwdOD2aOn1tpi8T3ZBWrbdcOFUNlS-GIzRANfax8313ozTo6JUpA9lPUQL8uXPFr4uOBkgJtTUqjDqXs6JYj6g6nJHx1vzjjPiFBYBxi0YP6_HyDx-2ahNJOzPW42RyPVY1fGjiEChP9jLEh-0r; QN1=000094002eb446aaefd0a93e; QN300=organic; QN99=2627; QunarGlobal=10.71.248.178_569353a2_183303e437e_68a2|1662965768482; _i=VInJOydfjXwCAtgxYAxzawHcaxJq; QN601=1a803b2cc1b557b4162a4147fd643b5a; QN48=000085002f1046aaf048503e; fid=6cb42d17-45dd-4318-b400-78a1f99be4ee; QN57=16629657813530.4356569196640583; qunar-assist={%22version%22:%2220211215173359.925%22%2C%22show%22:false%2C%22audio%22:false%2C%22speed%22:%22middle%22%2C%22zomm%22:1%2C%22cursor%22:false%2C%22pointer%22:false%2C%22bigtext%22:false%2C%22overead%22:false%2C%22readscreen%22:false%2C%22theme%22:%22default%22}; QN205=organic; QN277=organic; csrfToken=XI6nCyrptPhzWAbcnzGDcoDBL6C24KzY; QN67=16176%2C11077; QN269=F55CF0A05A9111ED9EEBFA163E3ACAE1; Hm_lvt_15577700f8ecddb1a927813c81166ade=1667381845; QN71=NjAuMTY4LjMyLjIxOuWQiOiCpTox; _vi=nDBBaBHqDknMeod-tpNS9vDXExwKtiL-rXh11i7ql-cOhGXAFkG7WSw9I_uTOO_6eOjY06eZT_bjAkN_Cmyxe_Oqkeh4GJ0dvVMGkE3H-vNX6fh563hImSU5bjOoWwWHPX6MPno7i9D2ekpuohGfYJWn9AWSoEBuZHlk7DzkLN4_; QN63=%E9%BB%84%E5%B1%B1; QN267=02201073991c670627; ariaDefaultTheme=undefined; QN58=1667394399732%7C1667394678113%7C5; Hm_lpvt_15577700f8ecddb1a927813c81166ade=1667394678; QN271=df7b3d16-6ea5-4d89-b2d0-c35365a1bea6; JSESSIONID=A8B874EBA013EF0F5CEF63E6D8B1F2F8; __qt=v1%7CVTJGc2RHVmtYMSsza1RaOW1vMGNoNTF2M0F3ZDllc3gwLy9sRnJDV1RXNmpIaXIzeUp1TE9xcWd3RzI3QlNpUzNocEhRK0dVdGRUVDd5NnZNUHZjK0xrdXBpOFRRZ0sxdkpPVC9pQytlWUtIaEZETWUwTm1LanR3MkJoKytkSG9kMnJUR21WVHdoYmxhR0oxcDRrY2xUTkFIYm10ejU4RVBad255aWNYczdVPQ%3D%3D%7C1667395490262%7CVTJGc2RHVmtYMTlKazZxejBNOFhmU1l0czlVUlZRWU44ZTBEL1ZKdzJFZjZhMzljTjJQQnBmUFVQSm1ja3hEK1c3VFFESWZIWVBhcXBEYWlrM2lxRGc9PQ%3D%3D%7CVTJGc2RHVmtYMTlxUkFJNXVhbUhwRk9lQjJ2MU4vSUZEQzAwUEZGZmVZZkZhVjdYMUx2a2w1ZlJwb3R1ZFdsb0tOamo1bXJ5dkZTN2dCUmp1c0NiZlI0TXpaYXVWa2Y2RVN0bDNGYmw4VXJJRXg2Rnc0WEVUQzFVTzlHYlhncEV1Wll3NTUrdjQxTjdDMFpqbFRmeDFqS1JwbHlYN00zT1pkNkwzNUZBMkR6LzRzWnczSUp3ei9SREZQS3dhWll6ZmdTcUVacm16U2laelVEZkZiUGZvRlhTVTk2Zmt2RW1sVUYwUG5WL1l2clRCL1cvNmg5d0JqSkxRVlFQSVVVcHpkQnV2Q2RVRkN0M1JJTUlKZHRWY09MUU42WXE1NW1qajVxYWhKcGNUZnF4MVFTSjVDa0pMamcxd1I5U3dGUFhGTFlmRHpHSTF0cnFUUWhDNkRvay84YTlYQXk3RGUzMy9LQ0hPeFdHRjdWNXlMVmZaVnhKaVdCcktMelRJQTBrODc4dDVOMDVjS2tjQ1N4RGhyNVZiOS9jK2NPWWhuWGNIRjA3czFqSmF5NFdsSzQvcWZ0OVRBSSt6SDg1WFZ4VEQ0UzNpTzVVQlRCR1lNcEdYYW1naTdPNGUwVklqdVU2MHJWaHBiVDVRY1BKS2t0WEsvczBVaWYzZHhMSmt3RGVjUFUrQzVCd1Fsa3ZIUGZEWFlKS3R3cVFPblBLQ0tYeEJxVkZ1ejhodWNteisxR1dHY1dQL3BUVDVrQysxckZxSmRRRFRNbk9yam1RdmkzZUhrdUdpMm5zZENsaWZMZnl0RGM2NDRzT3grdmhCc2I4U3licW1DM1ZVbDdQaEZRcEpGWjF3ZHcwSmZUcDhHdUtJa3BBYjVRRHhhWXJDelI3WmdpNnYyS25nVnVmWkcrczBuR3gybjJucDR5ODRTQmNndFRMc0dnemVGeE1vckxSMXYvVFNqbW9EMEU0Z2tNcUFvSUlaUHFUdFZoNXk0WXp3dWFlRjFiM3E3NDBaUWJrNnV1ZTVsUnI3MGlzS2d0WmV6dkdxYjBFUXo4dEFzYlhTeUxhT2IyWmxMNW5BRXlDbU9pSWtRLytYb0FZbkhWY1c3ZUluVTk0M3NNR0hRWFNwTm50NFE9PQ%3D%3D'
     }
     r = requests.get(url=url, params=param, headers=headers)
     j = r.json()
     c = j['data']
     num = [i for i in range(len(c['commentList']))]
     print(a)
     for x in num:
         if (c['commentList'][x]['content'] != "用户未点评，系统默认好评。"):
          print(c['commentList'][x]['score'])
          print(c['commentList'][x]['content'])
          i += 1
          str1 = 'A' + str(i)
          str2 = 'B' + str(i)
          sheet[str1].value = c['commentList'][x]['score']
          sheet[str2].value = c['commentList'][x]['content']
          workbook.save(filename ='qunaer1.xlsx')
     time.sleep(1)
     a+=1
