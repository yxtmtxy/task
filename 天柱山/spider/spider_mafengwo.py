import re
import time
import requests
from openpyxl import load_workbook

workbook = load_workbook(filename ='mafengwo.xlsx')
sheet = workbook.active
#评论内容所在的url，？后面是get请求需要的参数内容
comment_url='https://pagelet.mafengwo.cn/poi/pagelet/poiCommentListApi?'

requests_headers={
    'Referer': 'https://www.mafengwo.cn/poi/5429399.html',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36',
    'cookie': 'mfw_uuid=62089c6b-2077-87b0-a732-6f9faceac05b; __jsluid_h=b937905b567edec9b26addf473e1a64c; uva=s%3A286%3A%22a%3A4%3A%7Bs%3A13%3A%22host_pre_time%22%3Bs%3A10%3A%222022-02-13%22%3Bs%3A2%3A%22lt%22%3Bi%3A1644731499%3Bs%3A10%3A%22last_refer%22%3Bs%3A159%3A%22https%3A%2F%2Fwww.baidu.com%2Flink%3Furl%3DeIraOCi_glTMiOIL-L3efUke6rZNuAfljia-xMXelVymC0GMnyF38HWF2J56gZE48fQNqGfI-vrw27ehgH1oh_%26wd%3D%26eqid%3De5787834000cb10e0000000362089c66%22%3Bs%3A5%3A%22rhost%22%3Bs%3A13%3A%22www.baidu.com%22%3B%7D%22%3B; __mfwurd=a%3A3%3A%7Bs%3A6%3A%22f_time%22%3Bi%3A1644731499%3Bs%3A9%3A%22f_rdomain%22%3Bs%3A13%3A%22www.baidu.com%22%3Bs%3A6%3A%22f_host%22%3Bs%3A3%3A%22www%22%3B%7D; __mfwuuid=62089c6b-2077-87b0-a732-6f9faceac05b; __jsluid_s=1759c33b644913e1655bf2e2a4d55584; _r=bing; _rp=a%3A2%3A%7Bs%3A1%3A%22p%22%3Bs%3A12%3A%22cn.bing.com%2F%22%3Bs%3A1%3A%22t%22%3Bi%3A1662298265%3B%7D; __jsl_clearance_s=1664091582.316|0|BAJlZjum1rO6%2BV3gC96GwNUAL1s%3D; PHPSESSID=brc32a6umh5o6afaq2l9ch7jh5; oad_n=a%3A3%3A%7Bs%3A3%3A%22oid%22%3Bi%3A1029%3Bs%3A2%3A%22dm%22%3Bs%3A15%3A%22www.mafengwo.cn%22%3Bs%3A2%3A%22ft%22%3Bs%3A19%3A%222022-09-25+15%3A39%3A44%22%3B%7D; __mfwc=direct; __mfwa=1644731498952.22615.13.1663147508426.1664091583578; __mfwlv=1664091583; __mfwvn=11; bottom_ad_status=0; __omc_chl=; __omc_r=; __mfwb=50ee5646936f.3.direct; __mfwlt=1664091592'
}#请求头
i=165
for num in range(11,15):
    requests_data={
        'params': '{"poi_id":"5429399","page":"%d","just_comment":1}' % (num)   #经过测试只需要用params参数就能爬取内容
        }
    response =requests.get(url=comment_url,headers=requests_headers,params=requests_data)
    if 200==response.status_code:
        page = response.content.decode('unicode-escape', 'ignore').encode('utf-8', 'ignore').decode('utf-8')#爬取页面并且解码
        page = page.replace('\\/', '/')#将\/转换成/
        #星级列表
        star_pattern = r'<span class="s-star s-star(\d)"></span>'
        star_list = re.compile(star_pattern).findall(page)
        #评论列表
        comment_pattern = r'<p class="rev-txt">([\s\S]*?)</p>'
        comment_list = re.compile(comment_pattern).findall(page)
        for num in range(0, len(star_list)):
            #星级评分
            star = star_list[num]
            #评论内容，处理一些标签和符号
            comment = comment_list[num]
            comment = str(comment).replace('&nbsp;', '')
            comment = comment.replace('<br>', '')
            comment = comment.replace('<br />', '')
            print(star+"\t"+comment)
            i += 1
            str1 = 'A' + str(i)
            str2 = 'B' + str(i)
            sheet[str1].value = star
            sheet[str2].value = comment
            workbook.save(filename='mafengwo.xlsx')
    else:
        print("爬取失败")

