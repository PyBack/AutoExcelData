# -*- coding: utf-8 -*-

import os
import time
import pywinauto
from pywinauto.application import Application

import auto_helper as helper

os.chdir("c:/Users/Administrator/Downloads")

pwa_app = pywinauto.application.Application()
app_md = Application().start("C:/Users/Administrator/AppData/Local/atom/atom.exe")
time.sleeep(8)
windows = pywinauto.application.findwindows.find_windows(title_re="Welcome")
w_atom_handle = windows[0]
window_editor = pwa_app.window_(handle=w_atom_handle)
window_editor.Wait('ready', timeout=30)
window_editor.SetFocus()
window_editor.Maximize()
window_editor.Restore()
window_editor.SetFocus()

product_type = 'DLS'
target_series_list = [2998, 2999]
# target_series_list = list(xrange(23616, 23616+1))
# target_series_list.append(22289)
# target_series_list.append(22290)

for series_cout in target_series_list:

    if product_type == 'ELS':
        filename = u"삼성증권 제%d회 주가연계증권(공모) 상품설명서_최종.html" % series_cout
    elif product_type == 'DLS':
        filename = u"삼성증권 제%d회 기타파생결합증권(공모) 상품설명서_최종.html" % series_cout
    # app = Application().start("notepad.exe")
    app = Application().start(u"notepad.exe " + filename)
    app.window_().Wait('ready', timeout=30)

    app.window_().SetFocus()
    app.window_().Click()
    helper.pressHoldRelease('ctrl', 'a')
    helper.copy()

    # app_md = Application().start("C:\Users\Administrator\AppData\Local\Programs\MarkdonwPad 2\MarkdownPad2.exe")
    # app_md = Application().start("C:\Program Files (x86)\MarkdownPad 2\MarkdownPad2.exe")
    # app_md.window_().Wait('ready', timeout=30)

    window_editor.SetFocus()
    window_editor.Maximize()
    window_editor.Restore()
    window_editor.ClickInput(coords=(600, 100))  # avoid intersect with notepad.exe window
    helper.pressHoldRelease('ctrl', 'n')
    helper.paste()
    helper.pressHoldRelease('ctrl', 's')
    time.sleep(2)
    helper.typer('%s.html' % series_cout)
    helper.press('enter')
    time.sleep(3)
    helper.pressHoldRelease('ctrl', 'w')

    app.window_().Close()

