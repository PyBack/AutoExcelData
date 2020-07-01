# -*- coding: utf-8 -*-

import os
import time
import logging
import datetime as dt
import pywinauto
from pywinauto.application import Application

import auto_helper as helper
from handler_front import handle_front

logger = logger.getLogger('AutoOTC.download')
# logger.setLevel(logging.DEBUG)

# create file handler which logs even debug messages
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
# logger.addHandler(ch)

# os.chdir("C:/Users/Administrator/Downloads")


def download_statement_from_front(window_front=None, trade_id_list=[]):
    if window_front is None:
        logger.info("no handle of front...")
        return

    if not window_front.Exists():
        logger.error('no handler of front...')
        return

    window_front.SetFocus()

    window_front.ClickInput(coords=(80, 340))  # URL Editor
    clipboard.copy('http://front.samsungsecurities.local/instrument/document/index/F190621-00001')
    helper.paste()
    helper.press('enter')
    time.sleep(0.5)

window_front = handle_front()
time.sleep(3)
trade_id_list = ['F190626-00006',
                 'F190626-00005',
                 'F190626-00004',
                 ]

trade_id_list = ["F200618-%05d" % i for i in xrange(5, 16+1)]
trade_id_list = trade_id_list + ["F200618-%05d" % i for i in xrange(2, 3+1)]

for trade_id in trade_id_list:
    window_front.SetFocus()
    logger.info("downloading %s" % trade_id)
    window_front.ClickInput(coords=(340, 50))  # URL Editor
    # clipboard.copy('http://45.249.2.37/instrument/document/index/%s' % trade_id)
    # clipboard.copy('http://45.249.2.37/Document/Trade/%s' % trade_id)
    clipboard.copy('http://front.samsungsecurities.local/Document/Trade/Detail/%s' % trade_id)
    helper.press('backspace')
    helper.paste()
    helper.press('enter')
    time.sleep(5)
    window_front.ClickInput(coords=(890, 500))   # download button final
    time.sleep(2)

logger.info("======== END download %d ==========" % len(trade_id_list))
