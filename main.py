# -*- coding: utf-8 -*-

from __future__ import print_function

import os
import time
import datetime as dt
import logging
import win32api
import win32con
import pandas as pd
import clipboard
import pywinauto
import pywinauto.application as application
import auto_helper as helper
import xlwings as xw

import excel_control
from handler_hnet import handle_hnet
from read_data_file import read_aitas_std_price

logger = logging.getLogger('AutoReport')
logger.setLevel(logging.DEBUG)

# create file handler which logs even debug messages
# fh = logging.FileHandler('AutoReport.log')
fh = logging.handlers.RotatingFileHandler('AutoReport.log', maxBytes=104857, backupCount=3)
fh.setLevel(logging.DEBUG)

# create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)

# create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(name)s %(message)s')
fh.setFormatter(formatter)
ch.setFormatter(formatter)

# add the handler to logger
logger.addHandler(fh)
# logger.addHandler(ch)

VK_CODE = helper.VK_CODE

fundcode_lst = [
                '1356004',
                '1356010',
                ]

fund_column_lst = [1, 19]


fund_column_dict = dict(zip(fundcode_lst, fund_column_lst))

# 아이타스 펀드코드 조회용
fund_asset_manage_code_dict = dict()
fund_asset_manage_code_dict['1356020'] = u'8714'
fund_asset_manage_code_dict['1358001'] = u'D633'
fund_asset_manage_code_dict['1358006'] = u'D640'
fund_asset_manage_code_dict['1358011'] = u'D646'

fund_data_dict = {}

now_dt = dt.datetime.now()
strdate = now_dt.strftime("%Y%m%d")
strdate1 = now_dt.strftime("%Y-%m-%d")

# date_lst = list()

excel_file_name = u'수익률보고서_%s.xlsx' % strdate[2:]


def get_date_lst():
    global excel_file_name
    sheet_name = u'국내 기준가'

    wb = xw.apps[0].books[excel_file_name]
    sht = wb.sheets[sheet_name]
    sht.activate()

    df = sht.range((9, 1), (255, 1)).options(pd.DataFrame).value
    date_lst = pd.to_datetime(df.index)

    return date_lst


def get_backward_bizdate(date, date_lst):
    while not (date in date_lst):
        date = date - pd.DateOffset(1)
        date = date.to_pydatetime()

    return date


def retreive_stdprice_data_from_hnet(window_hnet=None):
    # now_dt = dt.datetime.now()

    if window_hnet is None:
        logger.error('no handle of hnet ...')
        return

    if not window_hnet.Exists():
        logger.error('no handle of hent ...')
        return

    window_hnet.SetFocus()

    # 30126 기준가정보
    sub_window_title = u'30126 \ud380\ub4dc\uae30\uc900\uac00\uaca9\uc815\ubcf4'
    sub_window = window_hent[sub_window_title]

    if not sub_window.Exists():
        window_hnet.ClickInput(coords=(70, 70))     # Editor (# of sub_window)
        clipboard.copy('30126')
        helper.paste()
        # helper.press('enter')

    sub_window.Maximize()
    sub_window.Restore()
    sub_window.SetFocus()
    sub_window.ClickInput(coords=(70, 15))
    helper.press('2')

    for fundcode in fundcode_lst:
        sub_window.ClickInput(coords=(70, 55))      # 펀드코드 Edit
        # clipboard.copy('1356004')
        clipboard.copy(fundcode)
        helper.paste()

        # sub_window.ClickInput(coords=(725, 95))   # Click 조회
        helper.press('enter')

        # Copy Data

        sub_window.RightClickInput(coords=(70, 205))
        helper.press('up_arrow')
        time.sleep(0.3)
        helper.press('up_arrow')
        time.sleep(0.3)
        helper.press('enter')
        time.sleep(0.3)

        text = clipboard.paste()
        text_line_lst = text.split('\n')
        text_line_encode = text_line_lst[2].encode('utf-8')
        text_line_encode = text_line_encode.replace('/', '-')
        text_line_encode_lst = text_line_encode.split('\t')
        fund_code_info = "%s %s %s %s %s" % (fundcode,
                                             text_line_encode_lst[0],
                                             text_line_encode_lst[1],
                                             text_line_encode_lst[5],
                                             text_line_encode_lst[8],
                                             )
        logger.info(fund_code_info)

        data = text_line_lst[2]
        data = data.replace('\r', '')
        data = data.replace('/', '-')
        data_lst = data.split('\t')

        fund_data_dict[fundcode] = data_lst

    pass


def retrieve_buysell_data_from_hnet(window_hnet=None):

    if window_hnet is None:
        logger.info('no handle of hnet...')
        return

    if not window_hnet.Exists():
        logger.error('no handle of hnet...')
        return

    window_hnet.SetFocus()

    # 30137
    sub_window_title = u'30137 \uc124\uc815\ud574\uc9c0\ud604\ud669'
    sub_window = window_hnet[sub_window_title]

    if not sub_window.Exists():
        window_hnet.ClickInput(coords=(70, 70)) # Editor (# of sub_window)
        clipboard.copy('30137')
        helper.paste()
        # helper.press('enter')

    sub_window.Maximize()
    sub_window.Restore()
    sub_window.SetFocus()

    sub_window.ClickInput(coords=(40,15)) # 자동조회
    sub_window.ClickInput(coords=(120,15)) # 클래스통합
    sub_window.ClickInput(coords=(120,125)) # 펀드유형3
    helper.press('2')

    helper.press('enter')

    logger.info('waiting data retreive...')
    time.sleep(30) # FIXME: check retreive done
    logger.info('stop waiting 30 sec')
    # Copy Data

    sub_window.RightclickInpu(coords=(70, 200))
    helper.press('up_arrow')
    time.sleep(0.3)
    helper.press('up_arrow')
    time.sleep(0.3)
    helper.press('enter')
    time.sleep(0.3)

    text = clipboard.paste()
    return textpre
    pass


def excel_process_domestic_stdprice():
    now_dt = dt.datetime.now()
    strdate = now_dt.strftime("%Y%m%d")
    strdate1 = now_dt.strftime("%Y-%m-%d")

    # strdate = '20180105'
    # strdate1 = '2018-01-05'

    global excel_file_name
    sheet_name = u'국내 기준가'

    excel_control.insert_row(excel_file_name, sheet_name, 9)
    df_select = read_aitas_std_price(strdate)

    wb = xw.apps[0].books[excel_file_name]
    wb.activate()
    sht_domestic_std_price = wb.sheets[sheet_name]
    sht_domestic_std_price.activate()

    for fundcode in fundcode_lst:

        sht_domestic_std_price.select()
        xw.apps[0].visible = True

        column_index = fund_column_dict[fundcode]
        rng = sht_domestic_std_price.range((3, column_index+1))
        fund_name = rng.value
        rng = sht_domestic_std_price.range((9, column_index))
        rng.select()
        xw.apps[0].visible = True

        time.sleep(1)

        fund_asset_manage_code = fund_asset_manage_code_dict.get(fundcode, '')
        fund_info_text = "%s %s %s" % (fundcode,
                                       fund_asset_manage_code,
                                       fund_name)
        logger.info(fund_info_text)

        if len(df_select.columns) > 0 and len(df_select) > 0:
            df_target = df_select[df_select[u'펀드코드'] == fund_asset_manage_code]
        else:
            df_target = pd.DataFrame()

        if fund_asset_manage_code != '' and len(df_target) == 1:
            adj_std_price = float(df_target[u'수정기준가'])
            fund_info_lst = ['aitas_price',
                             strdate1,
                             str(fundcode),
                             str(fund_asset_manage_code),
                             '%.3f' % adj_std_price
                             ]
            fund_info_text = ' '.join(fund_info_lst)
            logger.info(fund_info_text)

            data_lst = [strdate1, adj_std_price]
            fund_data_dict[fundcode] = data_lst

        elif fund_asset_manage_code != '' and len(df_target) > 1:
            column_index = fund_column_dict[fundcode]
            excel_control.insert_range(excel_file_name, sheet_name, [[9, column_index], [9, column_index+16]], len(df_target)-1)
            df_target_data = df_target[[u'수정기준가']]
            df_target_data.index = pd.to_datetime(df_target_data.index) + pd.DateOffset(days=1)
            df_targt_data.index = df_target_data.index.astype(str)
            df_target_data.sort_index(ascending=True)

            logger.info(str(df_target_data))

            start_row = 9

            for i in range(len(df_target)):
                holi_date = df_target_data.index[i]
                adj_price = float(df_target_data.iloc[i])
                cell_range = sht_domestic_std_price.range((start_row + i, column_index))
                cell_range.value = holi_date
                cell_range = sht_domestic_std_price.range((start_row + i, column_index + 1))
                cell_range.value = adj_price

            continue

        elif fund_asset_manage_code != '' and len(df_target) == 0:
            data_lst = fund_data_dict[fundcode]
            std_price = data_lst[1]
            std_price_prev = data_lst[3]
            fund_info_lst = ['no aitas_price',
                             strdate1,
                             str(fundcode),
                             str(fund_asset_manage_code),
                             str(std_price),
                             str(std_price_prev),
                             ]
            fund_info_text = ' '.join(fund_info_lst)
            logger.info(fund_info_text)

            column_index = fund_column_dict[fundcode]

            rng1 = sht_domestic_std_price.range((10, column_index+1))
            rng2 = sht_domestic_std_price.range((9, column_index+5))

            address1 = rng1.get_address(False, False)
            address2 = rng2.get_address(False, False)

            std_price = std_price.replace(',', '')
            std_price_prev = std_price_prev.replace(',', '')

            # formula_text1 = u"""=INDIRECT(%s&10)*(1+INDIRECT(%s&9))""" % (column_letter1, column_letter2)
            formula_text1 = u"=+(1+%s)*%s" %(address2, address1)
            formula_text2 = u"=+(1+%s)*%s" %(std_price, std_price_prev)
            logger.info(formula_text1)
            logger.info(formula_text2)
            data_lst = list()
            data_lst.append(strdate1)
            data_lst.append(formula_text1)
            data_lst.append('')
            data_lst.append('')
            data_lst.append('')
            data_lst.append(formula_text2)

            fund_data_dict[fundcode] = data_lst

            rng1 = sht_domestic_std_price.range((9, column_index + 1))
            rng2 = sht_domestic_std_price.range(address2)
            rng1 = rng1.api
            rng2 = rng2.api

            rng1.Font.Color = 15773696.0
            rng2.Font.Color = 15773696.0

        column_index = fund_column_dict[fundcode]
        data_lst = fund_data_dict[fundcode]
        data_lst[0] = strdate1

        # cell_range = sht_domestic_std_price.range((3, column_index))
        # fundcode_in_excel = str(cell_range.options(number=int).value)
        # print(fundcode_in_excel)

        cell_range = sht_domestic_std_price.range((9, column_index))
        cell_range.value = data_lst

    pass


def excel_process_buysell_pf(text):

    sht_name = u'설정해지(사모)'
    wb = xw.apps[0].books[excel_file_name]
    wb.activate()
    sht_buysell_pf = wb.sheets[sht_name]
    sht_buysell_pf.activate()
    sht_buysell_pf.select()

    sht_buysell_pf.range('A2').value = strdate1
    sht_buysell_pf.range('A4').expand().clear_contents()

    xw.apps[0].visible = True
    time.sleep(1)

    clipboard.copy(text)
    # excel_control.excel_process_buysell_pf_paste(excel_file_name, sht_name, 'A4')
    ws = sht_buysell_pf.api
    ws.Activate()
    rng = sht_buysell_pf.range('A4')
    rng.select()
    rng = rng.api
    rng.Select()
    xw.apps[0].visible = True
    rng.PasteSpecial()

    logger.info('paste excel_process_buysell_fund data')
    rng = sht_buysell_pf.range('A4')
    rng = rng.end('down')
    rng.select()

    xw.apps[0].visible = True

    time.sleep(3)
    pass


def excel_process_enter_date_monthly_ret():

    now_date = dt.date.today()
    month_count = (now_date.year - 2016 - 1) * 24 + (now_date.month - 1) - 9
    date_lst = get_date_lst()

    target_rownum = 41

    logger.info('month_count: %d' % month_count)

    sht_name = u'월별 수익률'
    wb = xw.apps[0].books[excel_file_name]
    wb.activate()
    sht_monthly_ret = wb.sheets[sht_name]
    sht_monthly_ret.activate()
    sht_monthly_ret.select()
    sht_monthly_ret.range('B4').select()

    df_1month = (now_date - pd.DateOffset(months=1)).to_pydatetime()
    df_3month = (now_date - pd.DateOffset(months=3)).to_pydatetime()
    df_6month = (now_date - pd.DateOffset(months=6)).to_pydatetime()

    df_1month = get_backward_bizdate(dt_1month, date_lst)
    df_3month = get_backward_bizdate(dt_3month, date_lst)
    df_6month = get_backward_bizdate(dt_6month, date_lst)

    str_dt_1month = dt_1month.strftime("%Y-%m-%d")
    str_dt_3month = dt_3month.strftime("%Y-%m-%d")
    str_dt_6month = dt_3month.strftime("%Y-%m-%d")

    logger.info('dt_1month: %s', str_dt_1month)
    logger.info('dt_3month: %s', str_dt_3month)
    logger.info('dt_3month: %s', str_dt_6month)

    sht_monthly_ret.range('B4').value = strdate1
    sht_monthly_ret.range('B4').select()
    xw.apps[0].visible = True
    time.sleep(1)
    sht_monthly_ret.range((target_rownum, month_count + 6)).select()
    sht_monthly_ret.range((target_rownum, month_count + 6)).value = str_dt_1month
    xw.apps[0].visible = True
    time.sleep(1)
    sht_monthly_ret.range((target_rownum, month_count + 7)).select()
    sht_monthly_ret.range((target_rownum, month_count + 7)).value = str_dt_3month
    xw.apps[0].visible = True
    time.sleep(1)
    sht_monthly_ret.range((target_rownum, month_count + 8)).select()
    sht_monthly_ret.range((target_rownum, month_count + 8)).value = str_dt_6month
    xw.apps[0].visible = True
    time.sleep(1)

    pass


def excel_process_insert_new_month_monthly_ret():
    now_date = dt.date.today()
    month_count = (now_date.year - 2016 - 1) * 24 + (now_date.month - 1) - 9
    date_lst = get_date_lst()

    target_rownum = 41

    logger.info('month_count: %d' % month_count)

    sht_name = u'월별 수익률'
    wb = xw.apps[0].books[excel_file_name]
    wb.activate()
    sht = wb.sheets[sht_name]
    sht.activate()

    sht_api = sht.api

    rng = sht_api.Columns(month_count + 5)
    rng.Select()
    rng.Insert()
    xw.apps[0].visible = True

    rng = sht_api.Columns(month_count + 4)
    rng.Select()
    rng.Copy()
    rng = sht_api.Columns(month_count + 5)
    rng.Select()
    rng.PasteSpecial()

    xw.apps[0].visible = True

    # update end of last month
    end_lastmonth = (now_date - pd.tseries.offsets.MonthEnd(1)).to_pydatatime()
    end_lastmonth = get_backward_bizdate(end_lastmonth, date_lst)

    str_end_lastmonth = end_lastmonth.strftime("%Y-%m-%d")

    sht.range((target_rownum, month_count + 5)).value = str_end_lastmonth
    sht.range((target_rownum, month_count + 5)).select()

    xw.apps[0].visible = True

    pass

def excel_process_monthly_return()
    now_date = dt.date.today()
    date_lst = get_date_lst()
    prev_bizdate = (now_date - pd.DateOffset(1)).to_pydatetime()
    prev_bizdate = get_backward_bizdate(prev_bizdate, date_lst)
    if now_date.day == 1:
        excel_process_insert_new_month_monthly_ret()
    elif prev_bizdate.month != now_date.month:
        excel_process_insert_new_month_monthly_ret()

    excel_process_enter_date_monthly_ret()
    pass

def excel_process_daily_report():

    sht_name = u'일보'
    wb = xw.apps[0].books[excel_file_name]
    wb.activate()
    sht_daily_report = wb.sheets[sht_name]
    sht_daily_report.range('T3').value = u"'기준일: %s" % strdate1
    sht_daily_report.activate()
    sht_daily_report.select()
    # logger.info('Range B2:T47 Copy')
    # excel_control.excel_process_copy_image_report(excel_file_name, sht_name, 'B2:T47')
    # wb.save()
    # logger.info("%s saved" % excel_file_name)
    pass

def main():
    logger.info("excel_file_name-> %s" % excel_file_name)
    window_hnet = handle_hnet()
    retreive_stdprice_data_from_hnet(window_hnet)
    excel_process_domestic_stdprice()
    text = retrieve_buysell_data_from_hnet(window_hnet)
    # text = clipboard.paste()
    excel_process_buysell_pf(text)

    excel_process_monthly_return()
    excel_process_daily_report()

if __name__ == "__main__":
    main()



