# -*- coding: utf-8 -*-

from __future__ import print_function


import time
import datetime as dt
import logging
import pandas as pd
import clipboard
import xlwings as xw

from read_data_file import read_otc_termination_file

logger = logging.getLogger('AutoOTC.Notice_HomePage')
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
logger.addHandler(ch)

pd.set_option('display.width', 180)


def get_target_product_data(excel_file_name='', strdate=''):
    if excel_file_name == '':
        excel_file_name = u'OTC상환리스트.xlsx'
    df = read_otc_termination_file(excel_file_name, strdate[:4] + "." + strdate[4:6])
    df = df.loc[strdate]

    if len(df) == 0:
        return df
    elif isinstance(df, pd.Series):
        df_new = pd.DataFrame()
        df_new = df_new.transpose()
        df = df_new.copy()

    target_df = df[(df[u'구분'] == 'ELS') | (df[u'구분'] == 'ELB') | \
                   (df[u'구분'] == 'DLS') | (df[u'구분'] == 'DLB')]
    target_df = target_df[(target_df[u'상환여부'] == u'만기상환') | (target_df[u'상환여부'] == u'조기상환')]
    target_df = target_df[(target_df[u'공사모'] == u'공모') | (target_df[u'공사모'] == u'공사모')]
    return target_df
    pass


def make_notice_msg(target_df, strdate=''):

    msg = "=== START make_notice_msg %d ===" % (len(target_df))
    logger.info(msg)

    now_dt = dt.datetime.now()
    year = now_dt.year
    month = int(strdate[4:6])
    day = int(strdate[-2:])

    f = open('notice_homepage.text', 'w')

    msg = u"[상환] %d년 %d월 %d일 기준 상환 확정 공모상품(%d건)" %(year, month, day, len(target_df))
    logger.info(msg)
    f.write((msg + u'\n' + u'\n').encode('ms949'))

    for i in range(len(target_df)):
        msg = u"%d. %s" % (i+1, target_df.iloc[i][u'상품명'])
        logger.info(msg)
        f.write((msg + u'\n').encode('ms949'))

        msg = u"%d. %s" % (i + 1, target_df.iloc[i][u'수익구조'])
        logger.info(msg)
        f.write((msg + u'\n').encode('ms949'))

        termination_value = target_df.iloc[i][u'상환예정단가'].astype(float) * 0.01
        msg = u" - 세전수익률: 원금의 %.2f%% 수준으로 상환" % termination_value
        logger.info(msg)
        f.write((msg + u'\n').encode('ms949'))

        exp_date = target_df.iloc[i][u'상환예정일']
        str_exp_date = u"%d. %0d. %0d" % (exp_date.year, exp_date.month, exp_date.day)
        msg = u" - 상환일자: %s 오전 11시지급 예정" % str_exp_date
        logger.info(msg)
        f.write((msg + u'\n').encode('ms949'))

    f.close()


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

    excel_file_name = u'OTC상환리스트.xlsx'
    target_df = get_target_product_data(excel_file_name, args.date)
    target_df = target_df[[u'종목코드', u'상품명', u'수익구조', u'기초자산1', u'기초자산2', u'기초자산3', u'상환여부', u'상환예정일', u'상환예정단가',]]
    logger.info(str(target_df))

    make_notice_msg(target_df, args.date)


if __name__ == "__main__":
    main()
