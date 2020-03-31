import requests
from bs4 import BeautifulSoup
import xlwt

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('sheet')

arrays =[]

def get_html(url):
    headers = {
        'User-Agent':'Mozilla/5.0(Macintosh; Intel Mac OS X 10_11_4)\
        AppleWebKit/537.36(KHTML, like Gecko) Chrome/52 .0.2743. 116 Safari/537.36'
    }

    response = requests.get(url,headers)
    html = response.text
    return html
soup = BeautifulSoup(get_html("http://hngcjs.hnjs.gov.cn/SiKuWeb/QiyeDetail.aspx?CorpName=河南省宛东建筑安装工程有限公司&CorpCode=91411328176722496J"),'lxml')
print(soup)
for tbody in soup.find_all("table"):
    trlist = tbody.find_all('tr')
    for tr in trlist:
        tblist = tr.find_all('td')
        for tb in tblist:
            if tb.string == None:
                pass
            else:
                if tb.string == '搜索':
                    pass
                else:
                    arrays.append(tb.string)

    # if i != 100:
    #     arrays.pop((i-1)*20*5)

for i in range(len(arrays)):
    print(arrays[i])
    worksheet.write(int(i/4), i%4, arrays[i])


workbook.save("/Users/jinyh/workspace/laowu1.xls")


