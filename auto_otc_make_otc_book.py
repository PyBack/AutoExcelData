# -*- coding: utf-8 -*-

from __future__ import print_function

import logging
import re
import datetime as dt
import pandas as pd

from read_data_file import read_otc_sales_file
from parsing_els_std_date import parsing_std_date

logger = logging.getLogger('AutoOTC.MakeOTCBook')
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


def convert_otc_book_format(df_data, strdate='', is_public=True):
    if is_public:
        f = open('otc_book_public.text', 'w')
    else:
        f = open('otc_book_private.text', 'w')

    for i in xrange(len(df_data)):
        row = df_data.iloc[i]
        product_name = row[u"상품명"]
        isin_code = row[u"종목코드"]
        deal_code = row[u"딜코드"]

        settle_currency = u"KRW"
        if product_name[-5:] == u"(USD)":
            settle_currency = product_name[-4:-1]

        sales_type1 = row[u"구분"]
        if sales_type1 in [u"ELT", u"DLT"]:
            continue

        sales_type2 = u''
        if is_public:
            sales_type2 = u"공모"
        else:
            sales_type2 = u"사모"

        sales_type3 = u''
        if row[u'회차(당/타발)'] == u"당사발행":
            sales_type3 = u"당발"
        elif row[u'회차(당/타발)'] == u"타사발행":
            sales_type3 = u"타발"

        issue_num = u'0'
        numbers = re.findall(u"\d+", product_name)
        if len(numbers) == 1:
            issue_num = numbers[0]
        pf_asset = sales_type1[:3] + issue_num
        online_exclusive = u''
        if is_public and row[u'비고'] == u"온라인전용":
            online_exclusive = u"O"

        btb_type = u"자체"
        initial_balance = row[u"판매금액"]
        initial_balance_fx = u''
        initial_balance_current = u''
        margin = row[u'마진(bp)']
        revenue = row[u'매출액']
        product_type = row[u'상품유형']
        underlying_type1 = row[u'기초자산유형1']
        underlying_type2 = row[u'기초자산유형2']
        underlying_asset_list = row[u'기초자산'].split(u"/")
        principle_protect = row[u"보장여부"]
        product_structure = row[u"상품설명"]

        results = parsing_std_date(int(issue_num))
        if sales_type1[:3] in [u'ELS', u'DLS']:
            midstrike_list = result[3]
        iniitial_price_list = results[4]
        if underlying_type2 == u"국내종목형":
            results[1] = results[1] + u'(3일)'

        if product_type in [u'(월지급식)', u'월지급SD(NoKI)']:
            product_structure_list = product_structure.split(' ')
            expire_n_period = product_structure_list[1]
            expire_n_period_list = expire_n_period.split('/')
            expirey = expire_n_period_list[0]
            periodic = expire_n_period_list[1]
            expirey = int(expirey[0])
            periodic = int(periodic[0])
            midstrike_list_new =  list()
            for i in xrange(len(midstrike_list)):
                if divmod(i + 1, periodic)[1] == 0:
                    midstrike_list_new.append(midstrike_list[i])

            if product_type in [u'SD', u'월지급SD', u'SD_멀티리자드']:
                product_structure_list = product_structure.split(' ')
                if sales_type1 == u"DLS":
                    ki_n_barrier = product_structure_list[3]
                else:
                    ki_n_barrier = product_structure_list[2]
                ki = ki_n_barrier.split('-')[0]
            elif product_type in [u'SD(NoKI)', u'월지급SD(NoKI)', u'SD_멀티리자드(NoKI)']:
                if sales_type1 == u"DLS":
                    ki_n_barrier = product_structure_list[3]
                else:
                    ki_n_barrier = product_structure_list[2]
                barrier = ki_n_barrier.split('-')[1]
                barrier = barrier.replace(u"(", "")
                barrier = barrier.replace(u")", "")
                barrier_list = barrier.split(u"")
                barrier_last = barrier_list[-1]
                ki = barrier_last
            else:
                ki = u""

            msg = "%s\t%s\t%s\t%s\t%s\t%s\t%s\t" % (product_name, isin_code, deal_code, settle_currency, sales_type1, sales_type2, sales_type3)
            msg = msg + "%s\t%s\t%s\t" % (pf_asset, online_exclusive, btb_type)
            msg = msg + "%s\t%s\t%s\t%s\t%s\t" % (initial_balance, initial_balance_fx, initial_balance_current, margin, revenue)
            msg = msg  + "%s\t%s\t%s\t" % (product_type, principle_protect, product_structure)
            if product_type not in [u'슈팅업', u'조기상환슈팅업']:
                msg = msg + "%s\t%s\t%s\t%s\t" % (results[0], results[1], results[2], results[1])
            elif product_type in [u'슈팅업', u'조기상환슈팅업']:
                msg = msg + "%s\t%s\t%s\t%s\t" % (results[0], results[1], results[2], results[1].strftime(u'%Y-%m-%d') + u'(3일')

            if product_type in [u'(월지급식)', u'월지급SD(NoKI)']:
                msg = msg + "\t".join([now_dt.strftime("%Y-%m-%d") for now_dt in midstrike_list_new])
                msg = msg + "\t".join(['' for i in xrange(37 - len(midstrike_list_new))])
            elif product_type == u"슈팅업":
                msg = msg + "\t".join(['' for i in xrange(36)])
            elif not (product_type in [u'KO Call & Put', u'하이파이브'] or underlying_type2 in [u"신용", u"신종자본증권", u"펀드"]):
                msg = msg + "\t".join([now_dt.strftime("%Y-%m-%d") for now_dt in midstrike_list])
                msg = msg + "\t".join(['' for i in xrange(37 - len(midstrike_list))])
            else:
                msg = msg + "\t".join(['' for i in xrange(36)])

            msg = msg + "%s\t%s\t" % (underlying_type1, underlying_type2)

            msg = msg + "\t".join(underlying_asset_list)
            msg = msg + "\t".join(['' for i in xrange(8 - len(underlying_asset_list))])

            msg = msg + "\t".join(["%f" % initial_price for initial_price in iniitial_price_list])
            msg = msg + "\t".join(['' for i in xrange(8 - len(underlying_asset_list))])

            msg = msg + "%s\t" % ki

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

    df_data = read_otc_sales_file(excel_file_name=u"OTC_판매현황_2020_ver1.3.xlsx", strdate=args.date)
    convert_otc_book_format(df_data, arg.date)


if __name__ == "__main__":
    main()

