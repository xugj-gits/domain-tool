#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Author:高效码农

import whois
import pprint
import xlsxwriter
from datetime import datetime
import concurrent.futures


# 批量读取文件中的域名
def read_file(filePath):
    with open(filePath, "r") as f:  # 打开文件
        data = f.readlines()  # 读取文件
        return data


# 通过某网站获取域名到期时间
def get_expiry_date(url):
    url_expiry_date_dict = {}
    try:
        whois_info = whois.whois(url.replace('\n', ''))
        endTime = whois_info.expiration_date
        print(whois_info)
    except whois.parser.PywhoisError as e:
        endTime = e
    url_expiry_date_dict['url'] = url.replace('\n', '')
    url_expiry_date_dict['endTime'] = endTime
    pprint.pprint(url_expiry_date_dict)
    url_expiry_date_list.append(url_expiry_date_dict)


def download_many(url_list):
    with concurrent.futures.ThreadPoolExecutor(max_workers=100) as executor:
        executor.map(get_expiry_date, url_list)


# 写入Excel文件
def write_excel(domain_list):
    # 创建一个新的文件
    with xlsxwriter.Workbook('host_ip1111.xlsx') as workbook:
        # 添加一个工作表
        worksheet = workbook.add_worksheet('域名信息')
        # 设置一个加粗的格式
        bold = workbook.add_format({"bold": True})
        # 分别设置一下 A 和 B 列的宽度
        worksheet.set_column('A:A', 50)
        worksheet.set_column('B:B', 15)
        # 先把表格的抬头写上，并设置字体加粗
        worksheet.write('A1', '域名', bold)
        worksheet.write('B1', '信息', bold)
        # 设置数据写入文件的初始行和列的索引位置
        row = 1
        col = 0
        for domain_ex_date in domain_list:
            url = domain_ex_date['url']
            endTime = domain_ex_date['endTime']
            currDate = datetime.today().date()
            try:
                endDate = endTime.date()
                diffDate = endDate - currDate
                if diffDate.days <= 7:
                    style = workbook.add_format({'font_color': "red"})
                else:
                    style = workbook.add_format({'font_color': "black"})
            except:
                style = workbook.add_format({'font_color': "red"})
            pprint.pprint(url + ': ' + str(endTime))
            worksheet.write(row, col, url, style)
            worksheet.write(row, col + 1, str(endTime), style)
            row += 1


urls = read_file('domain.txt')
url_expiry_date_list = []
download_many(urls)
write_excel(url_expiry_date_list)
