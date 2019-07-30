# -*- coding: utf-8 -*-

from __future__ import print_function


import time
import datetime as dt
import logging
import getpass
import pandas as pd
import clipboard
import auto_helper as helper
import xlwings as xw


from handler_hnet import handle_hnet
from read_data_file import read_otc_termination_file

logger = logging.getLogger('AutoOTC.Termination')
logger.setLevel(logging.DEBUG)

# create file handler whhich logs even debug messages
# fh = logging.FileHandler('AutoReport.log')
fh = logging.handlers.RotatingFileHandler('AutoOTC.log', maxBytes=104857, backupCount=3)
fh.setLevel(logging.DEBUG)

# create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)

# create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s [%(levelname)s %(name)s %(message)s')
fh.setFormatter(formatter)
ch.setFormatter(formatter)

# add the handler to logger
logger.addHandler(fh)
# logger.addHandler(ch)


excel_file_name = u'OTC상환리스트.xlsx'


def get_isin_code_lst(excel_file_name='', strdate=''):
    if excel_file_name == '':
        excel_file_name = u'OTC상환리스트.xlsx'
    df = read_otc_termination_file(excel_file_name, strdate[:4] + "." + strdate[4:6])
    df = df.loc[strdate]

    if len(df) == 0:
        return []
    elif isinstance(df, pd.Series):
        isin_code_lst = list()
        isin_code_lst.append(df[u'종목코드'])
        return isin_code_lst

    isin_code_lst = list(df[u'종목코드'])
    return isin_code_lst
    pass


def get_start_rownum(excel_file_name='', strdate=''):
    if excel_file_name == '':
        excel_file_name = u'OTC상환리스트.xlsx'
    df = read_otc_termination_file(excel_file_name, strdate[:4] + "." + strdate[4:6])
    df = df.reset_index()
    rownum = len(df[df[u'가격결정일'] < strdate])  # strdate: 전영업일
    return rownum
    pass


def get_prev_bzdate(excel_file_name='', strdate=''):
    if excel_file_name == '':
        excel_file_name == u'OTC상환리스트.xlsx'
    df = read_otc_termination_file(excel_file_name, strdate)
    df = df.reset_index()
    if strdate == '':
        now_dt = dt.datetime.now()
        strdate = now_dt.strftime("%Y%m%d")
    prev_bzdate = df[df[u'가격결정일'] < strdate].iloc[-1][u'가격결정일']
    str_prev_bzdate = prev_bzdate.strftime("%Y%m%d")
    return str_prev_bzdate
    pass


def manual_copy(sub_window, x_pos, y_pos):
    sub_window.RightClickInput(coords=(x_pos, y_pos))
    helper.press('down_arrow')
    helper.press('down_arrow')
    helper.press('down_arrow')
    helper.press('enter')
    pass


def query_otc_termination_data_from_hnet(window_hnet=None, isin_code_lst=[]):
    if window_hnet is None:
        logger.info('no handle of hent...')
        return

    if not window_hnet.Exists():
        logger.error('no handle of hnet...')
        return

    window_hnet.SetFocus()

    # 32802 파생결합증권상품정보
    sub_window_title = u'32802 파생결합증권상품정보'
    sub_window = window_hnet[sub_window_title]

    if not sub_window.Exists():
        window_hnet.ClickInput(coords=(70, 70))  # Editor (# of sub_window)
        clipboard.copy('32802')
        helper.paste()
        helper.press('enter')
        time.sleep(0.5)

    sub_window.Maximize()
    sub_window.Restore()
    sub_window.SetFocus()

    msg = '== START of query_otc_termination_data_from_hnet ==='
    logger.info(msg)

    termination_data_dict = dict()

    for isin_code in isin_code_lst:

        sub_window.ClickInput(coords=(90, 35))  # 종목코드
        clipboard.copy(isin_code[2:])
        helper.paste()

        # helper.press('enter')
        sub_window.ClickInput(coords=(775, 35))  # 조회
        time.sleep(0.1)

        sub_window.ClickInput(coords=(90, 15))  # 업무구분
        helper.press('down_arrow')
        helper.press('up_arrow')
        helper.press('up_arrow')
        helper.press('enter')

        sub_window.ClickInput(coords=(250, 275))  # 만기일
        manual_copy(sub_window, 250, 275)
        expire_date = clipboard.paste()
        if expire_date[:4] != "    ":
            expire_date = dt.datetime.strptime(expire_date, "%Y/%m/%d")
            str_expire_date = expire_date.strftime("%Y-%m-%d")
        else:
            str_expire_date = ""

        sub_window.ClickInput(coords=(700, 250))  # 상환예정일
        manual_copy(sub_window, 700, 250)
        termination_date = clipboard.paste()
        if termination_date[:4] != "    ":
            termination_date = dt.datetime.strptime(termination_date, "%Y/%m/%d")
            str_termination_date = termination_date.strftime("%Y-%m-%d")
        else:
            str_termination_date = "    -  -  "

        sub_window.ClickInput(coords=(700, 275))    # 상환예상단가
        helper.copy()
        termination_price = clipboard.paste()

        # ctrl = window_hnet['AfxWnd100u58']
        # ctrl.ClickInput()
        sub_window.ClickInput(coords=(100, 110))
        helper.copy()
        prdt_name = clipboard.paste()
        termination_type = u'미상환'
        if str_termination_date[:4] != "    ":
            if expire_date > termination_date:
                termination_type = u'조기상환'
            elif expire_date < termination_date:
                termination_type = u'만기상환'
            elif expire_date == termination_date:
                termination_type = u'카피에러'
            data_lst = [termination_type, str_termination_date, termination_price]
        else:
            data_lst = [termination_type, "", ""]

        msg = "%s %s %s %s %s %s" % (isin_code, prdt_name, str_expire_date,
                                     str_termination_date, termination_price, termination_type)
        logger.info(msg)

        termination_data_dict[isin_code] = data_lst
        sub_window.ClickInput(coords=(775, 35))  # 조회
        time.sleep(0.1)

    msg = "== END of query_otc_termination_data_from_hnet ==="
    logger.info(msg)
    logger.info('data count: %d' % len(termination_data_dict))
    return termination_data_dict
    pass


def excel_process_termination_data(isin_code_lst, termination_data_dict, rownum=3, sht_name=""):

    msg = '=== START excel_process_termination_data %d ===' % (len(isin_code_lst))
    logger.info(msg)

    now_dt = dt.datetime.now()
    if sht_name == "": sht_name = now_dt.strftime("%Y.%m")

    wb = xw.apps[0].books[excel_file_name]
    wb.activate()
    sht = wb.sheets[sht_name]
    sht.activate()
    sht.select()

    # rownum = sht.range('A3').current_region.last_cell.row

    count = 1
    for isin_code in isin_code_lst:
        data_lst = termination_data_dict[isin_code]
        if sht.range((rownum + count, 2)).value == isin_code:
            sht.range((rownum + count, 2)).value = data_lst
        else:
            logger.warn("isin_code isn't match rownum:%d isin_code:%s" % (rownum + count, isin_code))
        count += 1

    time.sleep(0.3)
    logger.info("=== END excel_process_termination_data ===")
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
    # pw = getpass.getpass("PWD: ")
    
    isin_code_lst = get_isin_code_lst(excel_file_name, args.date)
    logger.info('isin code count: %d' % len(isin_code_lst))
    
    # date_lst = ['20180212',
    #             '20180213',
    #             ]

    # date_rng = pd.bdate_range('2018-05-17', '2018-07-01')
    # date_lst = [d.strftime('%Y%m%d') for d in date_rng]

    window_hnet = handle_hnet()
    termination_data_dict = query_otc_termination_data_from_hnet(window_hnet, isin_code_lst)
    rownum = get_start_rownum(excel_file_name, args.date)
    excel_process_termination_data(isin_code_lst, termination_data_dict, rownum + 3, args.date[:4] + "." + args.date[4:6])
    pass


if __name__ == "__main__":
    main()
