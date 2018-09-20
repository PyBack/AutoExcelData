# -*- coding: utf-8 -*-

from __future__ import print_function

import time
import datetime as dt
import logging
import getpass
# import pandas as pd
import clipboard
import auto_helper as helper
import xlwings as xw

from handler_hnet import handle_hnet
from read_data_file import read_fund_code

logger = logging.getLogger('AutoReport.Fund_Day_BuySell')
logger.setLevel(logging.DEBUG)

# create file handler which logs even debug messages
# fh = logging.FileHandler('AutoReport.log')
fh = logging.handlers.RotatingFileHandler('AutoReport.log', maxByte=104857, backupCount=3)
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

excel_file_name = u'매매내역_hnet.xlsx'


def get_fund_code_lst(excel_file_name=''):
    if excel_file_name == '':
        excel_file_name = u'펀드코드.xlsx'
    df = read_fund_code(excel_file_name)
    if len(df) == 0:
        return []
    fund_code_lst = list(df[u'펀드코드'])
    return fund_code_lst
    pass


def retrieve_fund_day_buysell_data_from_hnet(window_hnet=None, fund_code_lst=[], buysell='B', pw='', strdate1=''):
    """
    :param window_hnet:
    :param fund_code_lst:
    :param buysell:
    :param pw:
    :param strdate1: '%Y%m%d'
    :return:
    """

    if window_hnet is None:
        logger.info('no handle of hnet...')
        return

    if not window_hnet.Exists():
        logger.error('no handle of hnet...')
        return

    window_hnet.SetFocus()

    # 30301 TR
    sub_window_title = u'30301 매매신청/결제현황'
    sub_window = window_hnet[sub_window_title]

    # if sub_window.Exists()
    #   sub_window.Close()

    if not sub_window.Exists():
        window_hnet.ClickInput(coords=(70, 70))  # Editor (# of sub_window)
        clipboard.copy('30301')
        helper.paste()
        # helper.press('enter')

    sub_window.Maximize()
    sub_window.Restore()
    sub_window.SetFocus()

    if strdate1 == '':
        now_dt = dt.datetime.now()
        strdate1 = now_dt.strftime('%Y%m%d')
        strdate = now_dt.strftime('%Y-%m-%d')
    else:
        strdate = strdate1[:4] + '-' + strdate1[4:6] + '-' + strdate1[-2:]

    msg = '=== START of retreive_fund_day_buysell_data_from_hnet %s %s ===' % (strdate1, buysell)
    logger.info(msg)

    sub_window.DoubleClickInput(coords=(90, 125))   # 처리일자
    helper.typer(strdate1)

    sub_window.ClickInput(coords=(620, 150))  # 비*밀&번$호
    clipboard.copy(pw)
    helper.paste()

    sub_window.ClickInput(coords=(90, 55))  # 매매구분
    if buysell == 'B':
        helper.press('1')  # 1. 매수 2. 매도
    elif buysell == 'S':
        helper.press('2')  # 1. 매수 2. 매도
    else:
        helper.press('1')  # 1. 매수 2. 매도

    line_data_lst = list()

    for fund_code in fund_code_lst:
        sub_window.ClickInput(coords=(90, 150))  # 펀드코드
        # clipboard.copy('1356016')
        clipboard.copy(fund_code)
        helper.paste()

        helper.press('enter')

        time.sleep(0.5)

        sub_window.RightClickInput(coords=(90, 250))  # Data
        helper.press('up_arrow')
        helper.press('up_arrow')
        helper.press('enter')

        text = clipboard.paste()
        text_line_lst = text.split('\n')
        line_count = len(text_line_lst)
        data_line_count = (line_count - 3) / 2

        for i in xrange(data_line_count):
            # text_line_encode = text_line_lst[i + 2].encode('utf-8')
            text_line_encode = text_line_lst[2*i + 2]
            text_line_encode = text_line_encode.replace('/', '-')
            text_line_encode_lst1 = text_line_encode.split('\t')

            # text_line_encode = text_line_lst[i + 3].encode('utf-8')
            text_line_encode = text_line_lst[2 * i + 3]
            text_line_encode = text_line_encode.replace('/', '-')
            text_line_encode_lst2 = text_line_encode.split('\t')

            data_lst = list()
            data_count = len(text_line_encode_lst1)

            if text_line_encode_lst1[0] == '':
                logger.info('no buysell fund: %s' % fund_code)
                continue
            else:
                data_lst.append(strdate)
                fund_name = text_line_encode_lst2[4]
                order_amount = text_line_encode_lst2[7]
                # print(fund_code, fund_name, str(order_amount))
                msg = str(fund_code) + ' '
                msg = msg + fund_name + ' '
                msg = msg + order_amount
                text_line_encode_lst1[2] = ''  # 계&좌*번$호
                logger.info(msg)

            for j in xrange(data_count):
                data_lst.append(text_line_encode_lst1[j])
                data_lst.append(text_line_encode_lst2[j])

            line_data_lst.append(data_lst)

    msg = '=== END of retreive_fund_day_buysell_data_from_hnet %s %s ===' % (strdate1, buysell)
    logger.info(msg)
    logger.info('data count: %d' % len(line_data_lst))

    return line_data_lst
    pass


def query_fund_buysell_able_data_from_hnet(window_hnet=None, fund_code_lst=[]):
    if window_hnet is None:
        logger.info('no handle of hnet...')
        return

    if not window_hnet.Exists():
        logger.error('no handle of hnet...')
        return

    window_hnet.SetFocus()

    # 30116 펀드정보등록
    sub_window_title = u'30116 펀드정보등록'
    sub_window = window_hnet[sub_window_title]

    # if sub_window.Exists():
    #     sub_window.Close()

    if not sub_window.Exists():
        window_hnet.ClickInput(coords=(70, 70))  # Editor (# of sub_window)
        clipboard.copy('30116')
        helper.paste()
        # helper.press('enter')

    sub_window.Maximize()
    sub_window.Restore()
    sub_window.SetFocus()

    # msg = '=== START of retreive_fund_day_buysell_data_from_hnet %s %s ===' % (strdate1, buysell)
    # logger.info(msg)

    # line_data_lst = list()

    for fund_code in fund_code_lst:

        sub_window.ClickInput(coords=(90, 150))  # 펀드코드
        # clipboard.copy('1356016')
        clipboard.copy(fund_code)
        helper.paste()

        helper.press('enter')
        print(fund_code)
        time.sleep(0.2)

    # msg = '=== END of retreive_fund_day_buysell_data_from_hnet %s %s ===' % (strdate1, buysell)
    # logger.info(msg)
    # logger.info('data count: %d' % len(line_data_lst))
    pass


def excel_process_buysell_pf(line_data_lst, buysell='B'):

    if buysell == 'B':
        sht_name = u'매수'
    elif buysell == 'S':
        sht_name = u'매도'
    else:
        sht_name = u'매수'

    msg = '=== START excel_process_buysell_pf %s %d ===' % (buysell, len(line_data_lst))
    logger.info(msg)

    wb = xw.apps[0].books[excel_file_name]
    wb.activate()
    sht = wb.sheets[sht_name]
    sht.activate()
    sht.select()

    rownum = sht.range('A4').current_region.last_cell.row
    # data_colnum = sht.range('A4').current_region.last_cell.column

    count = 1
    for data_lst in line_data_lst:
        sht.range((rownum + count, 1)).value = data_lst
        count += 1

    time.sleep(3)
    logger.info('=== END excel_process_buysell_pf ===')
    pass


def main():
    import argparse

    now_dt = dt.datetime.now()
    strdate = now_dt.strftime("%Y%m%d")
    parser = argparse.ArgumentParser()
    parser.add_argument('date',
                        type=lambda s: dt.datetime.strptime(s, "%Y%m%d").strftime("%Y%m%d"),
                        default=strdate,
                        help="Target Date",
                        nargs='?'
                        )

    args = parser.parse_args()
    logger.info("Target Date: %s" % args.date)
    pw = getpass.getpass()

    fund_code_lst = get_fund_code_lst()
    logger.info('fund code count: %d' % len(fund_code_lst))

    # date_lst = ['20180212',
    #             '20180213',
    #             '20180214',
    #             '20180219',
    #             '20180220',
    #             '20180221',
    #             '20180222',
    #             '20180223',
    #             ]

    window_hnet = handle_hnet()
    # query_fund_buysell_able_data_from_hnet(window_hnet, fund_code_lst)
    line_buy_data_lst = retrieve_fund_day_buysell_data_from_hnet(window_hnet, fund_code_lst, 'B', pw, args.date)
    line_sell_data_lst = retrieve_fund_day_buysell_data_from_hnet(window_hnet, fund_code_lst, 'S', pw, args.date)

    excel_process_buysell_pf(line_buy_data_lst, 'B')
    excel_process_buysell_pf(line_sell_data_lst, 'S')
    pass


if __name__ == "__main__":
    main()
