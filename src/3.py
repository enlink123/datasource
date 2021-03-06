import requests
from bs4 import BeautifulSoup
import xlwt
import re

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('航空港', cell_overwrite_ok=True)

arrays = []


# 郑州：410100
# 鹤壁：410600
# 开封：410200
# 洛阳：410300
# 平顶山：410400
# 安阳：410500
# 新乡：410700
# 焦作：410800
# 濮阳：410900
# 许昌：411000
# 漯河：411100
# 三门峡：411200
# 南阳：411300
# 商丘：411400
# 信阳：411500
# 周口：411600
# 驻马店：411700
# 济源：419001
# 巩义：410181
# 兰考：410225
# 汝州：410482
# 滑县：410526
# 林州：410581
# 长垣：410728
# 邓州：411381
# 永城：411481
# 固始：411525

# 鹿邑：411628
# 新蔡：411729
# 航空港：411800

def get_html(url, num):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36'
    }
    params = {
        "__EVENTTARGET": "AspNetPager2",
        "__EVENTARGUMENT": num,
        "__VIEWSTATE": "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",
        "__VIEWSTATEGENERATOR": "AB12D588",
        "__EVENTVALIDATION": "/wEdAAn4uMrNv1GN5xHYyNxjY+mlcyCwLNtjiGsUsD1klKe8mO/27pz3VRsK7NDdBcWPH1oVz/HNqEavkdJhHEQf0CC9QF5FF4kumzRC1Hm6gbSZLJSStlQIejt9Eiz2dXvmYMkdxEWiDJrToSQwV2qIPrDYCSAehAWh8K/R7+KNdzUNgpSYA6QBBGuDgl6JiXAF2qIuOK+UQVCRazqza48ZKfpBfBeHGajgPpe7w2qzYRJB7g==",
        "CretType": "建筑施工企业",
        "ry_reg_type": "411800"
    }
    response = requests.post(url, params, headers)
    html = response.text
    return html


def get_html_without_num(url):
    headers = {
        'User-Agent': 'Mozilla/5.0(Macintosh; Intel Mac OS X 10_11_4)\
        AppleWebKit/537.36(KHTML, like Gecko) Chrome/52 .0.2743. 116 Safari/537.36'
    }
    response = requests.get(url, headers)
    html = response.text
    return html


j = 0
writeTrIndex = 0
for i in range(1, 3):
    soup = BeautifulSoup(get_html("http://hngcjs.hnjs.gov.cn/company/list?page=" + str(i), 'lxml'))
    divContent = soup.find("div", id="tagContenth0")
    t = divContent.find("tr")
    trs = divContent.select('tr')  # 返回值类型为list
    print("pageNum=" + str(i))
    # print(trs)

    for trIndex, tr in enumerate(trs):
        # print(tr)
        tds = tr.select("td")
        for tdIndex, td in enumerate(tds):
            if tdIndex == 1:
                worksheet.write(writeTrIndex, 0, td.select("a")[0].get('href'))
        writeTrIndex = writeTrIndex + 1

    workbook.save("基本信息ID.xls")
