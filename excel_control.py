# -*- coding: utf-8 -*-

import logging
import win32com.client as win32com_client

excel = win32com_client.gencache.EnsureDispatch('Excel.Application')
logger = logging.getLoger('AutoReport.ExcelControl')

def insert_row(excel_file_name='', sheet_name='', range_cell_row=1):
    wb = None
    for i in range(1, excel.Workbooks.Count + 1):
        if excel.Workbooks.Item(i).Name == excel_file_name:
            wb = excel.Workbooks.Item(i)
            break
        if wb is None:
            logger.error('insert row-> no opened excel file %s' % excel_file_name)
            return
        
        logger.info('insert row-> catch %s file...' % wb.Name)
        
        wb.Activate()
        ws = wb.Worksheets(sheet_name)
        ws.Select()
        excel.Visible = True
        
        rng = ws.Range("%d:%d" %(range_cell_row, range_cell_row))
        rng.Select()
        rng.Insert()
        excel.Visible = True
        pass
    
def insert_range(excel_file_name = '', sheet_name='', range_cell=[[1,1], [1,1]], count=0):
    wb = None
    for i in range(1, excel.Workbooks.Count + 1):
        if excel.Workbooks.Item(i).Name == excel_file_name:
            wb = excel.Workbooks.Item(i)
            break
        if wb is None:
            logger.error('insert row-> no opened excel file %s' % excel_file_name)
            return
        
        logger.info('insert_range-> catch %s file...' % wb.Name)
        
        wb.Activate()
        ws = wb.Worksheets(sheet_name)
        cell1 = ws.Cells(range_cell[0][0], range_cell[0][1])
        cell2 = ws.Cells(range_cell[1][0], range_cell[1][1])
        ws.Range(cell1, cell2).Select()
        excel.Visible = True
        
        for i in range(count):
            ws.Range(cell1, cell2).Insert(Shift=win32com_client.constants.xlShiftDown)
            
def excel_process_buysell_pf_paste(excel_file_name='', sheet_name='', range_cell='A1'):
    wb = None
    for i in range(1, excel.Workbooks.Count + 1):
        if excel.Workbooks.Item(i).Name == excel_file_name:
            wb = excel.Workbooks.Item(i)
            break
        if wb is None:
            logger.info('excel_process_buysell_pf_paste-> no opened excel file %s' % excel_file_name)
            return
        
        logger.info('excel_process_buysell_pf_paste-> catch %s ...' % wb.Name)
        
        wb.Activate()
        ws = wb.Worksheets(sheet_name)
        ws.Select()
        ws.Activate()
        

