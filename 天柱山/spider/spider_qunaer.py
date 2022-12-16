import requests
import json
import re
import pandas as pd
import time
from openpyxl import load_workbook

workbook = load_workbook(filename ='qunaer.xlsx')
sheet = workbook.active

url='https://piao.qunar.com/ticket/detailLight/sightCommentList.json'
response = requests.get(url)
pagetext = response.text

a=1
i=0
while(a<575):
     param = {
               'sightId': '11077',
               'index': a,
               'page': a,
               'pageSize': '10',
               'tagType': '0'
     }
     headers = {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36',
          'cookie': 'QN1=00009100306c406a15b80410; QN300=s%3Dbing; QN99=7500; QunarGlobal=10.66.84.45_-2473e4f5_18022aefcd4_29a0|1649851582518; QN601=c0542051fb14f92be3f73c839c173c84; _i=VInJOQeP0A1UDBf3YqbRm4XE5I7q; QN48=000087802f10406a1658322c; fid=48b313af-3c1f-4cc2-8266-03e140f03ed9; QN57=16498516158980.009528324186312664; qunar-assist={%22version%22:%2220211215173359.925%22%2C%22show%22:false%2C%22audio%22:false%2C%22speed%22:%22middle%22%2C%22zomm%22:1%2C%22cursor%22:false%2C%22pointer%22:false%2C%22bigtext%22:false%2C%22overead%22:false%2C%22readscreen%22:false%2C%22theme%22:%22default%22}; QN205=s%3Dbing; QN277=s%3Dbing; csrfToken=kpcm2Yl5IM7xNslFGRaqDqVoI2DXJpNY; QN269=CF02A230326711ED90D7FA163E622C22; QN163=0; QN67=11077; _vi=4RwCMZR5YAh7S8XahgcATzhiLmOj9cg0xWPDSTr9a9DuGSQcLaLrkUr9dpv3JBs5EXkv0fLXpvirBLa2LEkJHAB8hlZ8jS3jFsQn-ULz-t8OAoI5tDMWMRhSVmoXrJZePztMOR5ibd8pIPu1lpvvTbHZfQLHykInTVvoDHKcWb12; HN1=v16ce97d8ed275eee7a65711e3055a4347; HN2=qunqgkknrlrns; ariaDefaultTheme=null; ctt_june=1654604625968##iK3wVRg%2BWhPwawPwa%3DGTa%3DkTa%3DkTaS2%3DXKv8ED38EKPsXPPnW%3DfGaRD%3DWKPNiK3siK3saKjOaS2OVKD%3DaSgmaUPwaUvt; ctf_june=1654604625968##iK3wWRPOawPwawPwasvAXSDsaRv%3DaR3naPETWR0RXs3AWRD%2BXsWhaKXsasPmiK3siK3saKjOaS2OVKD%3DaSgNWuPwaUvt; cs_june=624ea0965a890480fe0e4df2581d8b9247369f3eefccb39d0aa30c0f753c41440271b85c73ce096ce1aa9937ae11b3e9ddabce610868423a292da8a32bdd5c95b17c80df7eee7c02a9c1a6a5b97c11799cc6d750ae19772b014af625dcedc6ba5a737ae180251ef5be23400b098dd8ca; QN71="NjAuMTY4LjI1Mi4xNjE65ZCI6IKlOjE="; QN63=%E5%A4%A9%E6%9F%B1%E5%B1%B1; activityClose=1; QN243=4; QN267=107396326aa066512; QN58=1662969158723%7C1662969884132%7C8; QN271=299b345d-7a86-4012-ac26-984f712fbf8a; JSESSIONID=55366EF05D36825C1DB2419E2A4E27C0; __qt=v1%7CVTJGc2RHVmtYMSs5UnhOMHd1OWZjTnQ5UWlLMy9HVTNINkZscXBiSFB5UnpjMVFjT3NlTFl4ZXJyVEdiQ1JQODM5RzJnUEZEeEZLMWs5QU9RYXRTVndQeTJRQ3BMUW0xaGNuZXVOSkdsU1luZ2RGcDl4Mk81d0xSbkdQUklwRUdNRkhOMTkxM0R3QzBDYjdPVzB4dWxxSGZKNGtvNm8vK2M4WFh6dVd0dTAwPQ%3D%3D%7C1662969905134%7CVTJGc2RHVmtYMTlnWDZrcW13RkQyZ3g1cVhyVFEwQ1NQSzFOUXNhVWpIVGVSQ3M5V1ozRmdlU09jTjcvYzgzZG1YdG9QRlluWFlkdjZGYWYrQWpnaUE9PQ%3D%3D%7CVTJGc2RHVmtYMSt3QUlsZ3hUbEMvSEJyNksxWjVqVUtKZm5KSG5Ddk9vTDcwYnhRd0JaWnZIWW9IV3owYVpqelZ6c05TVUx3Z0pOL05vZHo3dmhlb2pINkgxOVQ2WWdjSTI3ZkRiNlh0a2tJTUtQM3NYZE1tWnNXejhUT2xTdWE0dUJvRTVLa3o3Q2xEZUpaeG9YUm5RbFF1RXZYVklCWkltUWgyYXZVVXpHUlExYjlpc2lQY2FoNlFhWkVwUVZBK00yUVRqeTdZV21Fai9zekhucENPUmptaHNTYi93aThmcGNOdmFPNXBxcWZYM0VZUFdhWDhMZy9rTTgxT214QjdJemdBYm5Mcm9SNTR5QnA4RXpEZ0pNZDd4YUZZOW15ZVRuSllzNTE3ZFdMdENUTzUza0NMSndQL0dkOGt5bnZ2WjFmMkE1WTAyMHdxaU92ZmNlNTc1RE1jTmpMZGdoY3hDMGlQUnI1d2c2Zi9uYkZNQW5DN09EL2tOSVN1eC80WVk4bjM0N2EvSEtielZxNTYvNTVyWVdUN0MxNDZDaVZTOE9PN2R2S1hRckltNm5paFlSRUs4clp1bFFDRlFtNzkwWThoYmRnUGpDSU1xRzVtQ3JmcCszeTZqUXpGRnFTU0xaRGNiSFdvbUFsQVRxOWpiZXpXNmhoRzY3WTA1U28xUjVZY0lsS1ppZWtLbGNOWHN0M3VrM05Rb3ZnYXBHTEwwMHFZUm80bDNBRmEyN0V5VWc3Q2ZEb0pGbFBYdjFDV0R2MzJOVHg3cjZnN0RJS1RlNEhRVnNSV0VHUTVjZFRKTW41OHZWRUMyM1Z3NEYwVzZwN21oTmt0ckFsOUFGdXlZNm5KeFRWTm4xNENoa2ZRcGdkM3hGU25RRGdGdSs5YVR3aEovbURTQWhmZHFUSFJLYThGQ1d3ZmsxaE9JZlp4M0FPOFMyQTVTTmV4MitTbDdjNjJ5K3hGeFdsckFtdmprekVtQjdjN05FTk5uVVdFeisrZ3QrT3dLWjIrb1BTVnhwNE92LzVoSURyQmtSbGJpSldaK2crMHR6cjFuNHBDaTl6aDR2WFE1UjcySWkzWWh2Nmdzb2M2VjVJaTRZOXBXZDJoN0JpdlMyQlU2MUxhY3A0WkE9PQ%3D%3D'
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
          workbook.save(filename ='qunaer.xlsx')
     time.sleep(1)
     a+=1
