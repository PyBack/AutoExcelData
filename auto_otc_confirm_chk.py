# -*- coding: utf-8 =*-

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


def get_confirm_isin_list_from_hnet(window_hnet=None):
    if window_hnet is None:
        logger.info('no handle of hnet')
        return

    if not window_hnet.Exists():
        logger.info('no handle of hnet')
        return

    window_hnet.SetFocus()

    # 30192 파생결합증권위험고지
    sub_window_title = u'32802 파생결합증권상품정보'
    sub_window = window_hnet[sub_window_title]

    if not sub_window.Exists():
        window_hnet.ClickInput(coords=(70, 70))  # Editor (# of sub_window)
        clipboard.copy('30192')
        helper.paste()
        helper.press('enter')
        time.sleep(0.5)

    sub_window.Maximize()
    sub_window.Restore()
    sub_window.SetFocus()

    msg = '== START of get_confirm_isin_code_list_from_hnet ==='
    logger.info(msg)

    sub_window.ClickInput(coords=(90, 15))  # 업무구분
    for i in xrange(6):
        helper.press('up_arrow')
    for i in xrange(3):
        helper.press('down_arrow')
    helper.press('enter')

    time.sleep(0.5)
    helper.press('enter')

    sub_window.RightClickInput(coords=(90, 140))
    helper.press('up_arrow')
    time.sleep(0.5)
    helper.press('up_arrow')
    time.sleep(0.5)
    helper.press('enter')
    time.sleep(0.5)

    data_table = clipboard.paste()
    data_table_rows = data_table.split("\n")

    isin_code_list = list()

    for row in data_table_rows:
        column_list = row.split("\t")
        if column_list[0] != u"상품코드" and len(column_list[0]) >= 10:
            isin_code_list.append(column_list[0])
            # print(column_list[0])

    logger.info("data load->isin_code cnt: %d" % len(isin_code_list))
    sub_window.Close()

    msg = "== END of get_confirm_isin_code_list_from_hnet ==="
    logger.info(msg)
    return isin_code_list


def get_total_settle_list_from_hnet(window_hent=None, strdate=None):
    if window_hent is None:
        logger.info('no handle of hent...')
        return

    if not window_hent.Exists():
        logger.info('no handle of hent...')
        return

    window_hent.SetFocus()

    # 66305 통합스케쥴내역1
    sub_window_title = u'66305 통합스케쥴내역1'
    sub_window = window_hent[sub_window_title]

    if sub_window.Exists():
        sub_window.Close()

    window_hent.ClickInput(coords=(70, 70))     # Editor (# of sub_window)
    clipboard.copy('66305')
    helper.paste()
    helper.press('enter')
    time.sleep(0.5)

    sub_window.Maximize()
    sub_window.Restore()
    sub_window.SetFocus()

    msg = '== START of get_total_settle_list_from_hnet ==='
    logger.info(msg)

    sub_window.DoubleClickInput(coords=(90, 15))    # 조회기간
    for i in xrange(2):
        for date_digit in strdate:
            helper.press(date_digit)

    sub_window.DoubleClickInput(coords=(90, 55))  # 종목종류
    for i in xrange(5):
        helper.press('down_arrow')
    for i in xrange(3):
        helper.press('up_arrow')
    helper.press('enter')

    sub_window.DoubleClickInput(coords=(700, 55))  # 일괄조회
    helper.press('enter')

    time.sleep(15)

    sub_window.ClickInput(coords=(90, 120))     # 자료 복사
    helper.press('up_arrow')
    time.sleep(1)
    helper.press('up_arrow')
    time.sleep(1)
    helper.press('enter')
    time.sleep(1)

    data = clipboard.paste()
    data = data.split("\r\n")
    new_data_lst = []
    for row in data:
        row_lst = row.split('\t')
        new_data_lst.append(row_lst)

    df_data = pd.DataFrame(new_data_lst)
    headers = df_data.iloc[0]
    df_data = pd.DataFrame(df_data.values[1:], columns=headers)
    # df_data.index = df_data[u'딜코드']
    df_data.index = df_data[df_data.columns[5]]

    sub_window.Close()

    msg = '=== END of get_total_settle_list_from_hnet ==='
    logger.info(msg)

    return df_data


def get_target_product_data(excel_file_name='', strdate='', term='상환'):
    if excel_file_name == '':
        excel_file_name = u'OTC상환리스트.xlsx'
    df = read_otc_termination_file(excel_file_name, strdate[:4] + "." + strdate[4:6])
    if not strdate in df.index:
        target_df = pd.DataFrame()
        # target_df = df.iloc[-2:].copy()
        return target_df
    df = df.loc[strdate]

    if len(df) == 0:
        return df
    elif isinstance(df, pd.Series):
        df_new = pd.DataFrame(df)
        df_new = df_new.transpose()
        df = df_new.copy()

    if term == '상환':
        target_df = df[(df[u'구분'] != 'ELT') & (df[u'구분'] != 'DLT')]
        target_df = target_df[(target_df[u'상환여부'] == u'만32805기상환') | (target_df[u'상환여부'] == u'조기상환')]
        return target_df
    elif term == '상환_ALL':
        target_df = df.copy()
        target_df = target_df[(target_df[u'상환여부'] == u'만기상환') | (target_df[u'상환여부'] == u'조기상환')]
        return target_df
    elif term == '미상환':
        target_df = df.copy()
        target_df = target_df[(target_df[u'상환여부'] == u'미상환')]
        return target_df
    else:
        df = pd.DataFrame()
        return df
    pass


def chk_in_isin_list(target_df, isin_code_list):
    # 파생결합증권위험고지 30192 4.수익확정
    msg = '=== START chk_in_isin_list %d ===' % (len(target_df))
    logger.info(msg)

    chk_in_list = []
    chk_in_count = 0

    for i in range(len(target_df)):
        exp_date = target_df.iloc[i][u'상환예정일']
        str_exp_date = u"%d-%0d-%d" % (exp_date.year, exp_date.month, exp_date.day)
        msg = u"%s %s %s %s %s " % (target_df.iloc[i][u'종목코드'],
                                    target_df.iloc[i][u'구분'],
                                    target_df.iloc[i][u'상환여부'],
                                    str_exp_date,
                                    target_df.iloc[i][u'수익구조'],
                                    )

        if target_df.iloc[i][u'종목코드'] in isin_code_list:
            chk_in_count += 1
            msg = msg + u" CHK_IN"
            chk_in_list.append(target_df.iloc[i][u'종목코드'])
        else:
            msg = msg + u"CHK_OUT"
        logger.info(msg)

    for isin in chk_in_list:
        isin_code_list.remove(msg)

    msg = 'CHK_IN: %s CHK_OUT: %s' % (chk_in_count, len(target_df) - chk_in_count)
    logger.info(msg)
    if chk_in_count > 0:
        msg = 'Must generate %d sms for termination H.Net #30192' % chk_in_count
        logger.warning(msg)
    msg = '=== END chk_in_isin_list ===='
    logger.info(msg)
    return isin_code_list
    pass


def chk_isin_in_salesteam(window_hnet, isin_code_list, df_data):

    # 32802 파생결합증권상품정보
    sub_windwow_title = u'32802 파생결합증권상품정보'
    sub_windwow = window_hnet[sub_windwow_title]

    if not sub_windwow.Exists():
        window_hnet.ClickInput(coords=(70, 70)) # Editor (# of sub_window)
        clipboard.copy('32802')
        helper.paste()
        helper.press('enter')
        time.sleep(0.5)

    sub_windwow.Maximize()
    sub_windwow.Restore()
    sub_windwow.SetFocus()

    msg = '== START of chk_isin_in_saleteam ==='
    logger.info(msg)

    df_data_sub = df_data[(df_data[u'Sales부서'] == u'PB') & (df_data[u'결제상태'] == u'최종확정')]
    df_data_early = df_data_sub[(df_data_sub[u'일자구분'] == 'OBS') & (df_data_sub[u'Sched.Type'] == u'의무중도')]
    df_data_mat = df_data_sub[(df_data_sub[u'일자구분'] == 'MAT')]

    df_data_sub = df_data[(df_data[u'Sales부서'] == u'PB') & (df_data[u'결제상태'] == u'미입력')]
    df_data_delay = df_data_sub[(df_data_sub[u'일자구분'] == 'OBS') & (df_data_sub[u'Sched.Type'] == u'의무중도')]

    logger.info('Sched Early Term-> %d' % len(df_data_early))
    logger.info('Sched Delay Term-> %d' % len(df_data_delay))
    logger.info('Sched MAT Term-> %d' % len(df_data_mat))

    for isin_code in isin_code_list:
        sub_windwow.ClickInput(coords=(90, 35))  # 종목코드
        clipboard.copy(isin_code[2:])
        helper.paste()

        # helper.press('enter')
        sub_windwow.ClickInput(coords=(775, 35))  # 조회
        time.sleep(0.5)

        sub_windwow.RightClickInput(coords=(110, 80))  # 딜코드
        helper.press('up_arrow')
        helper.press('up_arrow')
        helper.press('enter')

        deal_code = clipboard.paste()
        if deal_code in list(df_data.index):
            if len(df_data.loc[deal_code][u'Sales부서']) > 0:
                sales_team = df_data.loc[deal_code][u'Sales부서'][0]
                settle_state = df_date.loc[deal_code][u'결재상태'][0]
                sched_type = df_data.loc[deal_code][u'Sched.Type'][0]
            else:
                sales_team = df_data.loc[deal_code][u'Sales부서']
                settle_state = df_data.loc[deal_code][u'결재상태']
                sched_type = df_data.loc[deal_code][u'Sched.Type']
            msg = u"%s %s %s %s" % (isin_code, sales_team, settle_state, sched_type)
            logger.info(msg)
        else:
            logger.info("%s not in list" % deal_code)

        msg = '=== END of chk_isin_in_saleteam ==='
        logger.info(msg)


def chk_isin_in_schedul_list(window_hnet, target_df, df_data):
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

    msg = '== START of chk_isin_in_schedule_list ==='
    logger.info(msg)


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

    # date_lst = ['20180212',
    #             '20180213',
    #             ]

    # date_rng = pd.bdate_range('2018-05-17', '2018-07-01')
    # date_lst = [d.strftime('%Y%m%d') for d in date_rng]

    excel_file_name = u'OTC상환리스트.xlsx'
    window_hnet = handle_hnet()
    isin_code_list = get_confirm_isin_list_from_hnet(window_hnet)
    target_df = get_target_product_data(excel_file_name, args.date)
    if len(target_df) > 0:
        target_df = target_df[[u'종목코드', u'상품명', u'구분', u'수익구조', u'상환여부', u'상환예정일']]
        isin_code_list = chk_in_isin_list(target_df, isin_code_list)
        if len(isin_code_list) == 0:
            isin_code_list = list(target_df[u'종목코드'])
    print(isin_code_list)
    if len(isin_code_list) >= 0:
        df_data = get_total_settle_list_from_hnet(window_hnet, args.date)
        chk_isin_in_salesteam(window_hnet, isin_code_list, df_data)

    pass


if __name__ == "__main__":
    main()