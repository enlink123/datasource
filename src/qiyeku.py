import requests
from bs4 import BeautifulSoup
import xlwt
import json

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('sheet')

def read_score(jsonfile):
    with open(jsonfile, encoding='utf-8') as f:  # 将json文件转化为字典
        score_all = json.load(f)

    print(score_all)
    book = xlwt.Workbook()  # 创建excel文件
    sheet = book.add_sheet('sheet1')  # 创建一个表
    title = ['企业名称', '法定代表人', '注册资本', '统一社会信用代码', '企业类型', '注册所在地', '成立日期','备案施工专业']
    for col in range(len(title)):  # 存入第一行标题
        sheet.write(0, col, title[col])
    row = 1  # 定义行
    for k in score_all:
        data = score_all[k]  # data保存姓名和分数的list
        data.append(sum(data[1:4]))  # 倒数第二列加入总分
        data.append(sum(data[1:4]) / 3.0)  # 最后一列加入平均分
        data.insert(0, k)  # 第一列加入序号
        for index in range(len(data)):  # 依次写入每一行
            sheet.write(row, index, data[index])
        row += 1
    book.save('/Users/jinyh/workspace/qiyeku.xls')

read_score('/Users/jinyh/workspace/test.json')