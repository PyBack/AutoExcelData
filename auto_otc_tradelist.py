# -*- coding: utf-8 -*-

from __future__ import print_function

import datetime as dt
import logging
import xlwings as xw
import pandas as pd

logger = logging.getLogger('AutoOTC.TradeList')
logger.setLevel(logging.DEBUG)

# create file handler whhich logs even debug messages
fh = logging.FileHandler('AutoReport.log')
# fh = logging.handlers.RotatingFileHandler('AutoOTC.log', maxBytes=104857, backupCount=3)
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

pd.set_option('display.width', 240)

# Front 거래 관련 파일 읽기


def read_front_booktrade_file(excel_file_name="booktrade-20190430.xlsx", sales_team=u"상품개발팀", sales_channel=u"리테일 공모", counter_party=u"공모"):
    try:
        wb = xw.apps[0].books[excel_file_name]
    except:
        logging.info('no open excel file: %s' % excel_file_name)
        df = pd.DataFrame()
        return df

    sht = wb.sheets[u'Export']
    rownum = sht.range('D1').current_region.last_cell.row
    data_colnum = sht.range('D1').current_region.last_cell.column

    df = sht.range((1, 4), (rownum, data_colnum)).options(pd.DataFrame).value
    df = df.reset_index()
    df = df[(df[u'Sales부서'] == sales_team) & (df[u'판매채널세부분뷰'] == sales_channel)]
    if sales_team == u"신탁팀":
        df = df[df[u"최종거래상대방명"] == counter_party]
    df_output = df[df[u'Trader상태'] == u"확정"]
    df_output = df_output.sort_values(u'회차')

    return df_output


def read_private_trade_file(excel_file_name=u"사모설정.xlsx", strdate=''):
    try:
        wb = xw.apps[0].books[excel_file_name]
    except:
        logging.info('no open excel file: %s' % excel_file_name)
        df = pd.DataFrame()
        return df

    sht_name = strdate[:4] + '.' + strdate[4:6]
    sht = wb.sheets[sht_name]
    rownum = sht.range('A3').current_region.last_cell.row
    data_colnum = sht.range('A3').current_region.last_cell.column

    df = sht.range((3, 1), (rownum, data_colnum)).options(pd.DataFrame).value
    df = df.reset_index()
    df[u'거래일/홀딩일/확정예정일'] = pd.to_datetime(df[u'거래일/홀딩일/확정예정일'])
    df = df[(df[u'상품구분'] == u'신용') & (df[u'상태(확정/홀딩)'] == u'확정')]
    df_output = df[df[u'거래일/홀딩일/확정예정일'] == strdate]

    return df_output


def categorize_underlying_single_type(underlying_asset_single_name):
    underlying_type = u""
    if underlying_asset_single_name in [u"S&P500", u"EUROSTOXX50", u"HSCEI", u"KOSPI200", u"NIKKEI225", u"DAX"]:
        underlying_type = u"주가지수"
    elif underlying_asset_single_name in [u"WTI", u"Brent"]:
        underlying_type = u"Commodity"
    elif underlying_asset_single_name in [u"USDKRW"]:
        underlying_type = u"F/X"
    elif underlying_asset_single_name in [u"삼성전자", u"한국전력"]:
        underlying_type = u"국내종목"

    return underlying_type


def make_underlying_multi_type(underlying_num, underlying_type1, underlying_type2, underlying_type3):
    for i in xrange(underlying_num):
        if i == 0:
            underlying_type1
    pass


def make_report_equity_data(df, file_name='report_els_trade.text'):

    f = open(file_name, 'w')

    for i in xrange(len(df)):
        row = df.iloc[i]
        product_name = u""
        principle_protect = u""
        online_exclusive = u""

        if row[u"유형"] == "ELB":
            product_name = u"삼성증권 제" + row[u"회차"].astype(str)[:-2] + u"회 주가연계파생결합사채"
        elif row[u"유형"] == "ELS":
            product_name = u"삼성증권 제" + row[u"회차"].astype(str)[:-2] + u"회 주가연계증권"
        elif row[u"유형"] == "DLB":
            product_name = u"삼성증권 제" + row[u"회차"].astype(str)[:-2] + u"회 기타파생결합사채"
        elif row[u"유형"] == "DLS":
            product_name = u"삼성증권 제" + row[u"회차"].astype(str)[:-2] + u"회 기타파생결합증권"
        else:
            logger.error("not ELB or ELS")

        if row[u"결제통화"] != u"KRW":
            product_name = product_name + u"(" + row[u"결제통화"] + u")"

        product_type = row[u"유형"]
        product_structure_type = row[u"Template"]
        sales_team = u"상품개발"

        if row[u"거래ID"][0] == u"S":
            product_type = product_type + u"(POP)"

        if row[u"판매채널세부분류"] == u"리테일 신탁":
            product_type = product_type[:-1] + u"T"
            sales_team = u"신탁"

        if 0 < row[u"원금보장율(%)"] < 100:
            principle_protect = u"부분보장"
        elif row[u"원금보장율(%)"] >= 100:
            principle_protect = u"보장"
        elif row[u"원금보장율(%)"] == 0:
            principle_protect = u"비보장"

        if row[u"기초자산1"] == u"USD KRW 15:30 환율":
            row[u"기초자산1"] = "USDKRW"

        for i in xrange(3):
            if row[u"기초자산" + unicode(i+1)] == u"NKY225":
                row[u"기초자산" + unicode(i + 1)] = u"NIKKEI25"

        if row[u"기초자산3"] is not None:
            row[u"기초자산1"] = row[u"기초자산1"].replace(u"선물", "").strip()
            row[u"기초자산2"] = row[u"기초자산2"].replace(u"선물", "").strip()
            row[u"기초자산3"] = row[u"기초자산3"].replace(u"선물", "").strip()
            underlying_asset_name = row[u"기초자산1"] + u"/" + row[u"기초자산2"] + u"/" + row[u"기초자산3"]
            print(categorize_underlying_single_type(row[u"기초자산1"]), categorize_underlying_single_type(row[u"기초자산2"]), categorize_underlying_single_type(row[u"기초자산3"]))
        elif row[u"기초자산3"] is not None and row[u"기초자산2"] is not None:
            row[u"기초자산1"] = row[u"기초자산1"].replace(u"선물", "").strip()
            row[u"기초자산2"] = row[u"기초자산2"].replace(u"선물", "").strip()
            underlying_asset_name = row[u"기초자산1"] + u"/" + row[u"기초자산2"]
            print(categorize_underlying_single_type(row[u"기초자산1"]), categorize_underlying_single_type(row[u"기초자산2"]))
        else:
            row[u"기초자산1"] = row[u"기초자산1"].replace(u"선물", "").strip()
            underlying_asset_name = row[u"기초자산1"]
            print(categorize_underlying_single_type(row[u"기초자산1"]))
        underlying_asset_name = underlying_asset_name.strip()

        if row[u"온라인"] == 1:
            online_exclusive = u"온라인전용"
        else:
            online_exclusive = u""

        product_condition = u""
        if product_structure_type in [u"스텝다운", u"슈퍼스텝다운", u"월지급식NoKI", u"월지급식", u"멀티리자드", u"리자드"]:
            product_condition = underlying_asset_name + " "
            product_maturity = row[u"만기(년)"] + row[u"만기(월)"] / 12.0
            if row[u"만기(월)"] > 0:
                product_condition = product_condition + u"%.1f" % product_maturity + u"Y"
            else:
                product_condition = product_condition + u"%.0f" % product_maturity + u"Y"
            product_early_termination_freq = int(row[u'조기상환주기'][:-1])
            product_condition = product_condition + u"/%dM" % product_early_termination_freq
            if row[u'KI(%)'] > 0 and row[u'KI(%)'] != u'' and row[u'KI(%)'] - int(row[u'KI(%)']) == 0.0:
                product_condition = product_condition + u"%d-" % row[u'KI(%)']
            elif row[u'KI(%)'] > 0 and row[u'KI(%)'] != u'' and row[u'KI(%)'] - int(row[u'KI(%)']) != 0.0:
                product_condition = product_condition + u"%.1f-" % row[u'KI(%)']
            else:
                product_condition = product_condition + u" NOKI-"
                if product_structure_type == u"멀티리자드": product_structure_type = u"멀티리자드_NOKI"

            product_condition = product_condition + u"(%s)" % row[u'행사가(%)']
            product_condition = product_condition + u" 연 %.2f%%" % row[u'연쿠폰(%)']

            if product_structure_type in [u'월지급식', u'월지급식NoKI']:
                product_condition = product_condition + u" (월쿠폰베리어 %.1f, 월 %.2f%%)" % (row[u'조건1'], row[u'연쿠폰(%)'] / 12.0)

            if product_structure_type == u"리자드":
                msg = u" ("
                msg = msg + u"%d차 리자드베리어 %s%% 연쿠폰%s배" % (row[u'조건1'], row[u'조건2'], row[u'조건3'])
                msg = msg + u")"
                product_condition = product_condition + msg

            if product_structure_type in [u"멀티리자드", u"멀티리자드_NOKI"]:
                # print(row[u'조건1'].split(','), row[u'조건2'].split(','), row[u'조건3'].split(','))
                lizard_condition1_list = row[u'조건1'].split(',')
                lizard_condition2_list = row[u'조건2'].split(',')
                lizard_condition3_list = row[u'조건3'].split(',')
                msg = u" ("
                for i in xrange(len(lizard_condition1_list)):
                    msg = msg + u"%d차 리자드베리어 %s%% 연쿠폰*%s배" % (i+1, lizard_condition2_list[i], lizard_condition3_list[i])
                    msg = msg + u", "
                msg = msg + u")"
                product_condition = product_condition + msg

            if product_structure_type == u"스텝다운": product_structure_type = u"SD"
            elif product_structure_type == u"수퍼스텝다운": product_structure_type = u"SD_NOKI"
            elif product_structure_type == u"멀티리자드": product_structure_type = u"SD_멀티리자드"
            elif product_structure_type == u"멀티리자드_NOKI": product_structure_type = u"SD_멀티리자드(NoKI)"
            elif product_structure_type == u"월지급식": product_structure_type = u"월지급식SD"
            elif product_structure_type == u"월지급식NoKI": product_structure_type = u"월지급식SD(NoKI)"
            elif product_structure_type == u"양방향 낙아웃": product_structure_type = u"KO Call & Put"
            elif product_structure_type == u"NoKI원금보장": product_structure_type = u"하이파이브"

            office_name = row[u'상품조건']
            if office_name is None:
                office_name = ""
            if office_name[-2:] == u"지점":
                office_name = office_name[:-2]

            trade_margin = row[u"판매마진"].astype(float)

            msg = "%s\t%s\t%s\t%s\t%s\t\t\t%s" % (product_name, product_type, u"당사발행", product_structure_type, principle_protect, underlying_asset_name)
            msg = msg + "\t%s\t%s\t%.4f\t%.2f" % (row[u"청약시작"], row[u"거래일"], row[u"최초명목금액"] / 100000000.0, trade_margin)
            if office_name != "" and online_exclusive == u"":
                msg = msg + "\t\t%s\t%s" % (office_name, sales_team)
            else:
                msg = msg + "\t\t%s\t%s" % (online_exclusive, sales_team)
            msg = msg + "\t\t%s\t%s" % (product_condition, row[u"KRS코드"])
            logger.info(msg)
            f.write((msg + u'\n').encode('ms949'))

        f.close()


def make_report_credit_dls_data(df, file_name='report_credit_dls_trade.text'):
    f = open(file_name, 'w')

    for i in xrange(len(df)):
        row = df.iloc[i]
        product_name = u""
        principle_protect = u""
        online_exclusive = u""
        short_name_list = row[u"상품명_단축"].split(u' ')
        series_num = short_name_list[2][:-1]
        maturity = short_name_list[3]
        credit_sub_type = u' '
        if short_name_list[2] in [u'콜러블스탭업', u'선순위']:
            series_num = short_name_list[3][:-1]
            credit_sub_type = short_name_list[2]
            maturity = short_name_list[4]

        product_name = u"삼성증권 제" + series_num + u"회 기타파생결합증권"
        product_type = u"DLS"
        product_structure_type = u"기타"
        principle_protect = u"비보장"
        underlying_asset_name = short_name_list[0] + u" 신용"
        trade_date = row[u"거래일/홀딩일/확정예정일"].strftime("%Y-%m-%d")
        trade_amount = row[u"금액(억"].astype(float)
        trade_margin = row[u"마진(bp)"].astype(float)
        office_name = row[u"지점"]
        product_condition = short_name_list[0] + u" 신용연계 " + credit_sub_type + u" 만기" + maturity + u", 연" + u"%.2f%%" % (row[u"금리/쿠폰"].astype(float) * 100)
        product_condition = product_condition + u"*(이자기간/365)"
        sales_team = u"상품개발"
        krs_code = u""

        msg = "%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s" % (product_name, product_type, u"당사발행", product_structure_type, principle_protect, u"기타", u"신용", underlying_asset_name)
        msg = msg + "\t%s\t%s\t%.4f\t%.3f" % (trade_date, trade_date, trade_amount, trade_margin)
        msg = msg + "\t%s\t%s\t%s" % (online_exclusive, office_name, sales_team)
        msg = msg + "\t\t%s\t%s" % (product_condition, krs_code)
        logger.info(msg)
        f.write((msg + u"\n").encode('ms949'))

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

    excel_file_name = "booktrade-%s_Equity.xlsx" % args.date

    logger.info("======== ELS PUBLIC TRADE ========")
    df_output = read_front_booktrade_file(excel_file_name)
    # print(df_output)
    make_report_equity_data(df_output, 'report_els_public_trade.text')
    logger.info("==============================")

    logger.info("======== ELS PRIVATE TRADE ========")
    df_output = read_front_booktrade_file(excel_file_name, u"상품개발팀", u"리테일 사모")
    # print(df_output)
    make_report_equity_data(df_output, 'report_els_private_trade.text')
    logger.info("==============================")

    logger.info("======== ELT PUBLIC TRADE ========")
    df_output = read_front_booktrade_file(excel_file_name, u"신탁팀", u"리테일 신탁")
    # print(df_output)
    make_report_equity_data(df_output, 'report_elt_public_trade.text')
    logger.info("==============================")

    logger.info("======== ELT PRIVATE TRADE ========")
    df_output = read_front_booktrade_file(excel_file_name, u"상품개발팀", u"리테일 신탁", u"삼성증권신탁")
    # print(df_output)
    make_report_equity_data(df_output, 'report_elt_private_trade.text')
    logger.info("==============================")

    excel_file_name = "booktrade-%s_FICC.xlsx" % args.date

    logger.info("======== DLS PUBLIC TRADE ========")
    df_output = read_front_booktrade_file(excel_file_name)
    # print(df_output)
    make_report_equity_data(df_output, 'report_dls_public_trade.text')
    logger.info("==============================")

    logger.info("======== DLS PRIVATE TRADE ========")
    df_output = read_front_booktrade_file(excel_file_name, u"상품개발팀", u"리테일 사모")
    # print(df_output)
    make_report_equity_data(df_output, 'report_dls_private_trade.text')
    logger.info("==============================")

    logger.info("======== CREDIT DLS PRIVATE TRADE ========")
    df_output = read_private_trade_file(excel_file_name=u"사모설정.xlsx", strdate=args.date)
    make_report_credit_dls_data(df_output)
    logger.info("==============================")


if __name__ == "__main__":
    main()

