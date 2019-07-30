# -*- coding: utf-8 =*-

from __future__ import print_function


import time
import datetime as dt
import logging
import pandas as pd
import clipboard
import xlwings as xw
from xlwings.contants import SortOrder

from read_data_file import read_otc_termination_file

logger = logging.getLogger('AutoOTC.PB_Event')
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
        return []
    elif isinstance(df, pd.Series):
        df_new = pd.DataFrame(df)
        df_new = df_new.transpose()
        df = df_new.copy()

    # target_df = df[(df[u'구분'] == 'ELS') | \
    #                (df[u'구분'] == 'ELB') | \
    #                (df[u'구분'] == 'DLS') |]
    target_df = df[(df[u'구분'] != 'ELT') & (df[u'구분'] != 'DLT')]
    # target_df = target_df[target_df[u'상환여부'] == u'미상환']
    # target_df = target_df.sort(u'구분')
    return target_df
    pass


def make_excel_pb_event_msg(target_df, strdate=''):

    msg = '=== START excel_process_termination_data %d ===' % (len(isin_code_lst))
    logger.info(msg)

    now_dt = dt.datetime.now()
    excel_temp_file_name = 'pb_event_template.xls'

    wb = xw.apps[0].books[excel_temp_file_name]
    wb.activate()
    sht = wb.sheets[0]
    sht.activate()
    sht.select()
    sht.range("A1").select()
    rownum = sht.range('A2').current_region.last_cell.row
    colnum = sht.range('A2').current_region.last_cell.column
    sht.range(sht.cells(2, 1), sht.cells(rownum, colnum)).select()
    sht.range(sht.cells(2, 1), sht.cells(rownum, colnum)).clear_contents()
    for i in range(len(target_df)):
        sht.cells(i + 2, 1).value = target_df.iloc[i][u'종목코드']
        sht.cells(i + 2, 2).value = target_df.iloc[i][u'상품명']
        sht.cells(i + 2, 3).value = target_df.iloc[i][u'기초자산1']
        sht.cells(i + 2, 4).value = target_df.iloc[i][u'기초자산2']
        sht.cells(i + 2, 5).value = target_df.iloc[i][u'기초자산3']
        redemption_date = target_df.iloc[i][u'상환예정일']

        msg = strdate[:4] + u'년' + strdate[4:6] + u'월' + strdate[-2:] + u'일'
        msg_code = -1
        select_code = "-1"
        if target_df.iloc[i][u'상환여부'] == u"조기상환":
            msg = msg + u"조기상환 확정(%d/%d일 지급예정)_" % (redemption_date.mopnth, redemption_date.day)
            msg = msg + target_df.iloc[i][u'상품명']
            msg_code = 7
            select_code = "A"
            sht.cells(i + 2, 25).value = u"수익상환"
        elif target_df.iloc[i][u'상환여부'] == u"미상환":
            msg = msg + u"조기상환 순연_"
            msg = msg + target_df.iloc[i][u'상품명']
            msg_code = 2
            select_code = "B"
        if target_df.iloc[i][u'상환여부'] == u"만기상환":
            msg = msg + u"만기상환 확정(%d/%d일 지급예정)_" % (redemption_date.mopnth, redemption_date.day)
            msg = msg + target_df.iloc[i][u'상품명']
            msg_code = 8
            select_code = "A"
            sht.cells(i + 2, 25).value = u"수익상환"

        sht.cells(i + 2, 22).value = msg
        sht.cells(i + 2, 23).value = msg_code
        sht.cells(i + 2, 24).value = select_code

        logger.info("%s %d, %s" % (msg, msg_code, select_code))

    rownum = sht.range('A2').current_region.last_cell.row
    logger.info("rownum: %d" % rownum)
    wb.save()

    target_rng = sht.range((1, 1), (rownum, 25))
    # target_rng.select()
    # target_rng.api.Sort(Ket1=sht.range(sht.cells(2, 25), sht.cells(rownum, 25)).api, Order1=SortOrder.xlDescending)
    target_rng.api.AutoFilter
    target_rng.api.AutoFilter(Field=25, Criteria1=u"수익상환")
    target_path = u"C:/D_backup"
    msg = target_path + u'/%s_pb_event_수익상환.xls' % now_dt.strftime('%Y%m%d')
    logger.info(u"Excel Save: " + msg)
    wb.save(msg)

    # ToDo: sort & eliminate non-text 수익상환
    # msg = target_path + u'/%s_pb_event.xls' % now_dt.strftime('%Y%m%d')
    # logger.info(u"Excel Save: " + msg)

    logger.info('=== END make_excel_pb_event_statement ===')


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
    target_df = target_df[[u'종목코드', u'상품명', u'기초자산1', u'기초자산2', u'기초자산3', u'상환여부', u'상환예정일', u'구분']]
    logger.info(str(target_df))

    make_excel_pb_event_msg(target_df, args.date)


if __name__ == "__main__":
    main()
