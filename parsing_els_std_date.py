# -*- coding: utf-8 -*-

from __future__ import print_function

import re
import  datetime as dt
import urllib
from bs4 import BeautifulSoup


def convert_dt(strdate):
    year = int(strdate[:4])
    month = int(strdate[5:7])
    day = int(strdate[-3:-1])
    now_dt = dt.date(year, month, day)
    return now_dt


def parsing_std_date(series_count):

    file_name = urllib.urlopen("file:////c://Users/Administrator/Download/%d,html" % series_count)
    soup = BeautifulSoup(file_name, "html.parser")

    h1_lst = soup.findAll('h1')
    product_name = (h1_lst[0].text).split('\n')[0]
    product_type = (h1_lst[0].text).split('\n')[1].strip()
    # print ("#", product_name)
    # print (product_type)

    table_lst = soup.findAll('table', attrs={'class': 'MsoNormalTable'})

    table = table_lst[4]
    tr_lst = table.findAll('tr')

    issue_date = ''
    expire_date = ''
    initial_date = ''
    midstrike_date_list = list()

    for tr in tr_lst:
        td_lst = tr.findAll('td')
        td = td_lst[1]
        field_value = td.text
        field_value = field_value.replace('\n', '')
        field_value = field_value.replace('\r', '')

        td_0 = td_lst[0]
        field_name = td_0.text
        field_name = field_name.replace('\n', '')
        field_name = field_name.replace('\r', '')

        if field_name[:3] == u'(9)':
            # 발행일
            field_value = field_value.replace('\n', '')
            field_value = field_value[1:-1]
            now_dt = convert_dt(field_value)
            # print(field_name, now_dt)
            issue_date = now_dt
        elif field_name[:4] == u'(10)':
            # 만기일(예정)
            field_value = field_value.replace('\n', '')
            field_value = field_value[1:-1]
            now_dt = convert_dt(field_value)
            # print(field_name, now_dt)
            expire_date = now_dt

    table = table_lst[6]
    tr_lst = table.findAll('tr')

    for tr in tr_lst:
        td_lst = tr.findAll('td')
        td = td_lst[1]
        field_value = td.text
        field_value = field_value.replace('\n', '')
        field_value = field_value.replace('\r', '')

        td_0 = td_lst[0]
        field_name = td_0.text
        field_name = field_name.replace('\n', '')
        field_name = field_name.replace('\r', '')

        if field_name[:3] == u'(1)':
            # 최초기준가격
            # u'(1)\xa0\xa0 \ucd5c\ucd08\uae30\uc900\uac00\uaca9'
            # print(field_name, field_value)
            m = re.findall(r"(\[.*?\])", field_value)
            initial_price_list = list()
            for i in xrange(len(m)):
                if divmod(i, 2)[1] == 1:
                    initial_price = float(m[i].replace(',', '')[1:-1])
                    initial_price_list.append(initial_price)
                elif field_name[:3] == u'(2)':
                    # 최초기준가격결정일
                    field_value = td.text
                    field_value = field_value.split('\n')[1]
                    field_value = field_value.replace(' ', '')
                    field_value = field_value[1:-1]
                    now_dt = convert_dt(field_value)
                    # print(field_name, now_dt)
                    initial_date = now_dt
                elif field_name[:3] == u'(4)':
                    # 중간기준가격결정일
                    field_value = td.text
                    field_value_lst = field_value.split('\n')
                    for field_value in field_value_lst:
                        field_value = field_value.replace(' ', '')
                        field_value = field_value.replace('\r', '')
                        numbers = re.findall(u"^[0-9]]+", field_value)
                        if len(numbers) == 1:
                            now_dt = convert_dt(field_value[-11:])
                            # print(field_name, field_value[:3], now_dt)
                            midstrike_date_list.append(now_dt)

    return issue_date, expire_date, initial_date, midstrike_date_list, initial_price_list


if __name__ == "__main__":
    series_count = 23812
    results = parsing_std_date(series_count)
    print(results)
