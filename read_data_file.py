# -*- coding: utf-8 -*-

# import sys

import logging
import datetime as dt
import xlwings as xw
import pandas as pd

# reload(sy)
# sys.setdefaultencoding('UTF8')

logger = logging.getLogger('AutoReport.ReadAitasStdPrice')


def make_str(item):
    if type(item) == float:
        return unicode(int(item))
    else:
        return item

# 엑셀파일코드 내 펀드코드는 운용사펀드코드 Hnet #30116 참조


def read_aitas_std_price(strdate=''):
    excel_file_name = u'삼성증권_ 기준가격%s.xls' % (strdate)
    try:
        wb = xw.apps[0].books[excel_file_name]
    except:
        logging.info('no open excel file: %s' % excel_file_name)
        df = pd.DataFrame()
        return df

    sht = wb.sheets[u'sheet1']

    rownum = sht.range('A1').current_region.last_cell.row
    data_colnum = sht.range('A1').current_region.last_cell.column

    df = sht.range((1,1), (rownum, data_colnum)).options(pd.DataFrame).value
    df.index = df.index.astype(int)
    df.index = df.index.astype(str)
    df[u'펀드코드'] = df[u'펀드코드'].apply(make_str)
    df.index = pd.Series([item[:4] + '-' + itme[4:6] + '-' + item[-2:] for item in df.index])

    df[u'diff'] = abs(df[u'기준가'] - df[u'수정기준가'])
    df_select = df[(df[u'diff'] >0) & (df[u'펀드명'].str[-1] == u'A')]
    df_select = df_select.sort_index(ascending=False)
    logger.info(str(df_select))
    # data_text = str(df_select)
    # print(data_text)

    # wb.close()

    return df_select


# 엑셀파일코드 내 펀드코드는 사내관리 펀드코드

def read_fund_code(excel_file_name=u'펀드별 펀드코드.xlsx'):
    try:
        wb = xw.apps[0].books[excel_file_name]
    except:
        logging.info('no open excel file: %s' % excel_file_name)
        df = pd.DataFrame()
        return df

    sht = wb.sheets[u'Sheet1']
    rownum = sht.range('A1').current_region.last_cell.row
    data_colnum = sht.range('A1').current_region.last_cell.column

    df = sht.range((1,1), (rownum, data_colnum)).options(pd.DataFrame).value
    df[u'펀드코드'] = df[u'펀드코드'].apply(make_str)
    df = df.reset_index()

    return df


# OTC상환리스트 파일 읽기

def read_otc_termination_file(excel_file_name=u'OTC상환리스트.xlsx', sht_name=""):
    try:
        wb = xw.apps[0].books[excel_file_name]
    except:
        logging.info('no open excel file: %s' % excel_file_name)
        df = pd.DataFrame()
        return df

    now_dt = dt.datetime.now()
    if sht_name == "": sht_name = now_dt.strftime("%Y.%m")
    sht = wb.sheets[sht_name]
    sht.activate()
    rownum = sht.range('A3').current_region.last_cell.row
    data_colnum = sht.range('A3').current_region.last_cell.column

    df = sht.range((3, 1), (rownum, data_colnum)).options(pd.DataFrame).value

    return df


# OTC 원장 파일 읽기

def read_otc_book_file(excel_file_name="otc_book.xlsx"):
    try:
        wb = xw.apps[0].books[excel_file_name]
    except:
        print('no open excel file: %s' % excel_file_name)
        df = pd.DataFrame()
        return df

    sht = wb.sheets[u"DATA"]
    sht.activate()
    rownum = sht.range('A1').current_region.last_cell.row
    data_colunm = sht.range('A1').current_region.last_cell.column

    df = sht.range((1, 1), (rownum, data_colunm)).options(pd.DataFrame).value
    df = df.reset_index()

    return df


# OTC 판매현황 파일 읽기

def read_otc_sales_file(excel_file_name=u"OTC_판매현황_2020_ver1.3.xlsx", strdate='', is_pub_pvt='pub'):
    try:
        wb = xw.apps[0].books[excel_file_name]
    except:
        print('no open excel file: %s' % excel_file_name)
        df = pd.DataFrame()
        return df

    if strdate == '':
        now_dt = dt.datetime.now()
        strdate = now_dt.strftime("%Y%m%d")
    else:
        now_dt = dt.datetime.strptime(strdate, "%Y%m%d")

    sht = wb.sheets[u'%d월' % now_dt.month]
    sht.activate()

    data_colnum = 19
    if is_pub_pvt == 'pub':
        rownum = sht.range('A4').current_region.last_cell.row
        df = sht.range((4, 1), (rownum, data_colnum)).options(pd.DataFrame).value
    elif is_pub_pvt == 'pvt':
        rownum = sht.range('A4').end('down').end('down').end('down').row
        start_rownum = sht.range('A4').end('down').end('down').row + 1
        df = sht.range((start_rownum, 1), (rownum, data_colnum)).options(pd.DataFrame).value

    df = df.reset_index()

    df = df[df[u'판매 종료'] == strdate]

    return df


if __name__ == "__main__":
    import datetime as dt
    now_dt = dt.datetime.now()
    strdate = now_dt.strftime('%Y%m%d')
    # strdate = '20180102'
    # df_select = read_aitas_std_price(strdate)
    # df = read_fund_code(u'펀드별 펀드코드_180124.xlsx')
    # df = read_otc_termination_file(u'OTC상환리스트_201812.xlsx')
    df = read_otc_sales_file(strdate='20200617', is_pub_pvt='pvt')
    print(df)
    # print(df.loc[strdate])
