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
import auto_helpers as helpers
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



