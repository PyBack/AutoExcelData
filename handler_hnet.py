# -*- coding: utf-8 -*-

import logging
import pywinauto

logger = logging.getLogger('AutoReport.Handler_Hnet')


def handle_hnet():
    logger.info('=== START find Hnet handler ===')
    hnet_title = u'\uc0bc\uc131\uc99d\uad8c Honors Net'
    pwa_app = pywinauto.application.Application()
    w_hnet_handle = pywinauto.findwindows.find_window(title=hnet_title, class_name="POPHNET")[0]
    window_hnet = pwa_app.window_(handler=w_hnet_handle)
    window_hnet.Maximize()
    window_hnet.Restore()
    window_hnet.SetFocus()

    logger.info('=== END find Hnet handler ===')

    return window_hnet
    pass
