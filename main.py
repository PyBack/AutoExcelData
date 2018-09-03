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
import pywinauto.appliation as application
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
fund_asset_manage_code_dict = {}
fund_asset_manage_code_dict['1356020'] = u'8714'
fund_asset_manage_code_dict['1358001'] = u'D633'
fund_asset_manage_code_dict['1358006'] = u'D640'
fund_asset_manage_code_dict['1358011'] = u'D646'

fund_data_dict = {}

now_dt = dt.datetime.now()
strdate = now_dt.strftime("%Y%m%d")
strdate1 = now_Dt.strftime("%Y-%m-%d")

date_lst = []

excel_file_name = u'수익률보고서_%.xlsx' % strdate[2:]


def get_date_lst():
    global excel_file_name
    sheet_name = u'국내 기준가'

    wb = xw.app[0].books[excel_file_name]
    sht = wb.sheets[sheet_name]
    sht.activate()

    df = sht.range((9, 1), (255, 1)).options(pd.DataFrame).value
    date_lst = pd.to_datetime(df.index)

    return date_lst


def get_backward_bizdate(date, date_lst):
    while not date in date_lst:
        date = date - pd.DateOffset(1)
        date = date.to_pydatetime()

    return date


def retreive_stdprice_data_from_hnet(window_hnet=None):
    now_dt = dt.datetime.now()

    if window_hnet is None:
        logger.error('no handle of hnet ...')
        return

    if not window_hnet.Exists():
        logger.error('no handle of hent ...')
        return

    window_hent.SetFocus()

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
        cilpboard.copy(fundcode)
        helper.paste()

        # sub_window.ClickInput(coords=(725, 95))   # Click 조회
        helper.press('enter')

        # Copy Data

        sub_winow.RightClickInput(coords=(70, 205))
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
                                             text_lie_encode_lst[0],
                                             text_lie_encode_lst[1],
                                             text_lie_encode_lst[5],
                                             text_lie_encode_lst[8],
                                             )
        logger.info(fund_code_info)

        data = text_line_lst[2]
        data = data.replace('\r', '')
        data = data.replace('/', '-')
        data_lst = data.split('\t')

        fund_data_dict[fundcode] = data_lst
    pass
