import re
import execjs
import requests
import json
import hashlib
from requests.utils import add_dict_to_cookiejar
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from bs4 import BeautifulSoup
# 关闭ssl验证提示
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
header ={
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0',
    'Host' : 'www.mafengwo.cn'
}
session = requests.session() #使用session会一直携带上一次的cookies
url='https://www.mafengwo.cn/poi/5429399.html'
response = session.get(url, headers=header, verify=False) #直接访问得到JS代码
js_clearance = re.findall('cookie=(.*?);location', response.text)[0] #用正则表达式匹配出需要的部分

result = execjs.eval(js_clearance).split(';')[0].split('=')[1] #反混淆、分割出cookie的部分
add_dict_to_cookiejar(session.cookies, {'__jsl_clearance_s': result})  #将第一次访问的cookie添加进入session会话中
response = session.get(url, headers=header, verify=False) #带上更新后的cookie进行第二次访问
go = json.loads(re.findall(r'};go\((.*?)\)</script>', response.text)[0])
for i in range(len(go['chars'])):
    for j in range(len(go['chars'])):
        values = go['bts'][0] + go['chars'][i] + go['chars'][j] + go['bts'][1]
        if go['ha'] == 'md5':
            ha = hashlib.md5(values.encode()).hexdigest()
        elif go['ha'] == 'sha1':
            ha = hashlib.sha1(values.encode()).hexdigest()
        elif go['ha'] == 'sha256':
            ha = hashlib.sha256(values.encode()).hexdigest()
        if ha == go['ct']:
            __jsl_clearance_s = values

add_dict_to_cookiejar(session.cookies, {'__jsl_clearance_s' :__jsl_clearance_s})
response = session.get(url, headers=header, verify=False)




