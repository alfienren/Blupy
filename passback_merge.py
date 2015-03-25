from xlwings import Workbook, Range, Sheet
import pandas as pd
import numpy as np
import xlsxwriter
import win32com.client
import os

def passback_merge():

    wb = Workbook.caller()

    passback_folder = Range('Lookup', 'AB1').value
    sheet = Range('Lookup', 'AA1').value
    passback_templates = os.listdir(passback_folder)

    temp_to_merge = []
    for i in passback_templates:
        temp_to_merge.append(passback_folder + '\\' + str(i))

    merged_passback = pd.DataFrame(columns = ['Date', 'Campaign', 'Site', 'Placement', 'Spend',
                                              'Impressions', 'Clicks', 'Video Plays', '100% Video Completes'])

    for sheet in temp_to_merge:
        passback = pd.read_excel(sheet, 0, index_col = None, na_value=[0])
        merged_passback = merged_passback.append(passback)

    data = pd.DataFrame(pd.read_excel(sheet, 'data', index_col=None))

