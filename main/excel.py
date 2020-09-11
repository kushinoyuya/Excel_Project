import os
import glob
import openpyxl
import datetime

from pathlib import Path

def get_path():
    '''対象のエクセルファイルを選択する'''
    wb = openpyxl.load_workbook('./excel_project/sample.xlsx')
    ws = wb.worksheets[-1]
    wb.copy_worksheet(ws)
    ws.title = 'YYYYMM'
    wb.save('./excel_project/sample.xlsx')
