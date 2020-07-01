# -*- coding: utf-8 -*-

import logging
import pywinauto

logger = logging.getLogger('AutoReport.Handler_Front')


def handle_front():
    logger.info('=== START find Front handler ===')
    pwa_app = pywinauto.application.Application()
    chrome_title = u'\ubb38\uc11c - Google Chrome'
    # w_front_handle = pywinauto.findwindows.find_windows(title=chrome_title,class_name='Chrome_WidgetWin_1')[0]
    w_front_handle = pywinauto.findwindows.find_windows(class_name='Chrome_WidgetWin_1')[0]
    window_front = pwa_app.window_(handle=w_front_handle)
    window_front.Maximize()
    window_front.Restore()
    window_front.SetFocus()

    logger.info('=== END find Front handler ===')

    return window_front
    pass


if __name__ == "__main__":
    handle_front()