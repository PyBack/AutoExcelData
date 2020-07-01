# -*- coding: utf-8 -*-

import os
import win32com.client

# Create an instance of Word.Application
wordApp = win32com.client.Dispatch('Word.Application')

# Show the application
wordApp.Visible = True

product_type = 'DLB'
#
#
#
target_series_list = list(xrange(919, 919 + 1))

for i in target_series_list:
# for i in xrange(2498, 2500):
    # Open document in the application
    file_path = u"C:/Users/Administrator/Downloads"
    # filename = u"제[%d]회 주가 상품설명서_최종.html" %  i
    if product_type == 'ELS':
        file_name = u'삼성증권 제%d회 주가연계증권(공모) 상품설명서_최종.docx' % i
    elif product_type == 'DLS':
        file_name = u'삼성증권 제%d회 기타파생결합증권(공모) 상품설명서_최종.docx' % i
    elif product_type == 'ELB':
        file_name = u'삼성증권 제%d회 주가연계파생결합사채(공모) 상품설명서_최종.docx' % i
    elif product_type == 'DLB':
        file_name = u'삼성증권 제%d회 기타파생결합사채(공모) 상품설명서_최종.docx' % i
    docx_path = os.path.join('c:', os.sep, 'Users', 'Administrator', 'Downloads', file_name)
    doc = wordApp.Documents.Open(docx_path)
    # doc = None

    for i in range(1, wordApp.Documents.Count+1):
        if wordApp.Documents.Item(i).Name == file_name:
            doc = wordApp.Documents.Item(i)
            break
    # if docx is None:
    #     logger.error('insert row-> no opened excel file %s' % excel_file_name)
    #     return

    doc.Activate()
    text = doc.Range().Text
    print text[:40]

    if file_name[-3:] == "doc":
        docx_path = os.path.join('c:', os.sep, 'Users', 'Administrator', 'Downloads', file_name[:-4] + '.html')
    elif file_name[-4:] == "docx":
        docx_path = os.path.join('c:', os.sep, 'Users', 'Administrator', 'Downloads', file_name[:-5] + '.html')
    doc.SaveAs(FileName=docx_path, FileFormat=8)
    doc.Close()
