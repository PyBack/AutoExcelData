# -*- coding: utf-8 -*-
import pdb
import re
import urllib
from bs4 import BeautifulSoup

f = open("els_std_price.text", "w+")

one_star_sd = 24589

# target_series_list = [,]
# target_series_list = target_series_list + list(xrange(24439, 24441 + 1))
# target_series_list = target_series_list + list(xrange(24443, 24446 + 1))
# target_series_list = target_series_list + list(xrange(24448, 24452 + 1))

# for i in target_series_list:
for i in xrange(24503, 24605 + 1):

    series_count = i

    file_name = urllib.urlopen("file:////c://Users/Administrator/Download/%d,html" % series_count)
    soup = BeautifulSoup(file_name, "html.parser")

    h1_lst = soup.findAll('h1')
    product_name = (h1_lst[0].text).split('\n')[0]
    product_type = (h1_lst[0].text).split('\n')[1].strip()
    print "#", product_name
    print product_type

    table_lst = soup.findAll('table', attrs={'class': 'MsoNormalTable'})
    table = table_lst[6]


    tr_lst = table.findAll('tr')
    tr_lst[0]

    tr = tr_lst[0]
    td_lst = tr.findAll('td')
    td = td_lst[1]
    field_value = td.text
    field_value = field_value.replace('\n', '')
    field_value = field_value.replace('\r', '')

    td_0 = td_lst[0]
    field_name = td_0.text
    field_name = field_name.replace('\n', '')
    field_name = field_name.replace('\r', '')

    # print field_name, field_value
    # field_name == u'(1)    최초기준가격'

    table = table_lst[6]
    tr_lst = table.findAll('tr')

    monthly_field_name = ''
    monthly_field_value = ''

    lizard1_field_name = ''
    lizard1_field_value = ''

    lizard2_field_name = ''
    lizard2_field_value = ''

    for tr in tr_lst:
        td_lst = tr.findAll('td')

        td_0 = td_lst[0]
        field_name = td_0.text
        field_name = field_name.replace('\n', '')
        field_name = field_name.replace('\r', '')

        td_0 = td_lst[1]
        field_value = td_0.text
        field_value = field_value.replace('\r\n', '', 1)
        field_value = field_value.replace(' ', '')

        if field_name[:3] == u'(1)':
            # 최초기준가격
            # u'(1)\xa0\xa0 \ucd5c\ucd08\uae30\uc900\uac00\uaca9'
            print field_name, field_value
            m = re.findall(r"(\[.*?\])", field_value)
            initial_price = float(m[1].replace(',', ''))[1:-1]
        elif field_name[:3] == u"(5)":
            # 만기행사가격
            # u'(5)\xa0\xa0 \uc911\uac04\uae30\uc900\uac00\uac9 \uacb0\uc815\uc77c \ud589\uc0ac\uac00\uaca9'
            field_value = td.text
            field_value.strip()
            field_value = field_value.replace(' ', '')

            if product_type == u'(월지급식)' or product_type == u'(월지급식NoKI)':
                field_value = field_value.replace('\r\n', '', 1)
                monthly_field_name = field_name.replace(' ', '').replace(u'\xa0\xa0', '')
                monthly_field_value = field_value
            else:
                print field_name.replace(' ', '').replace(u'\xa0\xa0', '')
                print field_value.strip()
                print

                # f.write((field_name+u"\n").encode('ms949'))
                # f.write((field_value.strip()+u"\n").encode('ms949'))
                f.write("\n")

        elif field_name[:4] == u"(10)" and (product_type == u'(월지급식)' or product_type == u'(월지급식NoKI)'):
            # 월지급식 -> 중간기준가격 결정일 행사가격
            print field_name.replace(' ', '')
            field_value = td.text
            field_value = field_value.strip()
            field_value = field_value.replace(' ', '')
            print field_value
            print
            f.write((field_name+u"\n").encode('ms949'))
            f.write((field_value.strip()+u"\n").encode('ms949'))
        elif field_name[:4] == u"(6)" and product_type == u'(멀티리자드NoKI)':
            # print field_name, field_value
            lizard1_field_name = field_name.replace(' ', '').replace(u'\xa0\xa0', '')
            lizard1_field_value = field_value
        elif field_name[:4] == u"(7)" and product_type == u'(멀티리자드NoKI)':
            # print field_name, field_value
            lizard2_field_name = field_name.replace(' ', '').replace(u'\xa0\xa0', '')
            lizard2_field_value = field_value

    table = table_lst[8]
    tr_lst = table.findAll('tr')

    print u"최종기준가격 (만기행사가격, 하락한계가격)"

    for tr in tr_lst:
        td_lst = tr.findAll('td')

        td_0 = td_lst[0]
        field_name = td_0.text
        field_name = field_name.replace('\n', '')
        field_name = field_name.replace('\r', '')

        td_0 = td_lst[1]
        field_value = td_0.text
        field_value = field_value.replace('\r\n', '', 1)
        field_value = field_value.replace(' ', '')

        if product_type in [u'(조기상환슈팅업)', u'(슈팅업)']:
            # [final_strike]%
            # initial_price x final_strike
            break

        if field_name[:3] == u'(3)':
            m = re.findall(r"(\[.*?\])", field_value)
            final_price = float(m[1].replace(',', ''))[1:-1]
            final_ratio = final_price / initial_price * 100
            # print initial_price, final_price, final_ratio
            print "[%.1f]%%" % final_ratio
            print field_name.replace(u'\xa0\xa0', '')
            print field_value.strip()

            # f.write("[]%\n")
            # f.write((field_name+u"\n").encode('ms949'))
            # f.write((field_value.strip()+u"\n").encode('ms949'))

            if product_type == u'(월지급식)' or product_type == u'(월지급식NoKI)':
                m = re.findall(r"(\[.*?\])", monthly_field_value)
                monthly_price = float(m[1].replace(',', ''))[1:-1]
                monthly_ratio = monthly_price / initial_price * 100
                print "[%.1f]%%" % monthly_ratio
                print monthly_field_name.replace(' ', '').replace(u'\xa0\xa0', ''), monthly_field_value
                f.write("[]%\n")
                # f.write((monthly_field_name + u" " + monthly_field_value + u"\n").encode('ms949'))
            elif product_type.encode('utf-8') == '(멀티리자드NoKI)':
                m = re.findall(r"(\[.*?\])", lizard1_field_value)
                lizard_price = float(m[1].replace(',', ''))[1:-1]
                lizard_ratio = lizard_price / initial_price * 100
                print "[%.1f]%%" % lizard_price
                print lizard1_field_name.replace(' ', '').replace(u'\xa0\xa0', '')
                print lizard1_field_value.strip()

                m = re.findall(r"(\[.*?\])", lizard2_field_value)
                lizard_price = float(m[1].replace(',', ''))[1:-1]
                lizard_ratio = lizard_price / initial_price * 100
                print "[%.1f]%%" % lizard_price
                print lizard2_field_name.replace(' ', '').replace(u'\xa0\xa0', '')
                print lizard2_field_value.strip()

            # f.write("[]%\n")
            # f.write((lizard1_field_name + u"\n").encode('ms949'))
            # f.write((lizard1_field_name.strip() + u"\n").encode('ms949'))
            # f.write("[]%\n")
            # f.write((lizard2_field_name + u"\n").encode('ms949'))
            # f.write((lizard2_field_name.strip() + u"\n").encode('ms949'))

        elif field_name[:3] == u'(4)' and series_count == one_star_sd:
            # field_name[:9] != u'(4)\xa0\xa0 Worst':
            m = re.findall(r"(\[.*?\])", field_value)
            ki_price = float(m[1].replace(',', ''))[1:-1]
            ki_ratio = ki_price / initial_price * 100
            print "[%.1f]%%" % ki_ratio
            print field_name.replace(' ', '').replace(u'\xa0\xa0', ''), field_value

            # f.write("[]%\n")
            # f.write((field_value.strip()+u"\n").encode('ms949'))
        elif field_name[:3] == u'(5)' and series_count == one_star_sd and product_type[-5:] != u'NoKI)' and product_type != u'(수퍼스텝다운)':
            m = re.findall(r"(\[.*?\])", field_value)
            ki_price = float(m[1].replace(',', ''))[1:-1]
            ki_ratio = ki_price / initial_price * 100
            print "[%.1f]%%" % ki_ratio
            print field_name.replace(u'\xa0\xa0', '')
            print field_value.strip()

            f.write("[]%\n")
            # f.write((field_name+u"\n").encode('ms949'))
            # f.write((field_value.strip()+u"\n").encode('ms949'))

        print("")
        print("")

        f.write("\n")
        f.write("\n")

f.close()