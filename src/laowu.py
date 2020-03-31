import requests
from bs4 import BeautifulSoup
import xlwt

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('sheet')

arrays =[]

def get_html(url,num):
    headers = {
        'User-Agent':'Mozilla/5.0(Macintosh; Intel Mac OS X 10_11_4)\
        AppleWebKit/537.36(KHTML, like Gecko) Chrome/52 .0.2743. 116 Safari/537.36'
    }
    params ={
        "__EVENTTARGET":"ctl00$MainContent$AspNetPager2",
        "__EVENTARGUMENT":num,
        "__VIEWSTATE":"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",
        "__VIEWSTATEGENERATOR":"2D2EDCB8",
        "__EVENTVALIDATION":"/wEdAATl6l2SMXmWTtIoY/Mx0yNleQmEdRYi7qMbWkgmq5sjTk2v6HzSMfbZm//Kgj8po2iGF+rskpx0ux0IDXSXgMBa/HEsFKvuSBIfPW3vv12Crdvi3zycPiogroAWcUVkkb0="
        ,"CretType":"旧版劳务资质"
    }
    response = requests.post(url,params,headers)
    html = response.text
    return html
for i in range(121,156):
    soup = BeautifulSoup(get_html("http://hngcjs.hnjs.gov.cn/SiKuWeb/LaoWu.aspx",i),'lxml')

    print(i)
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


