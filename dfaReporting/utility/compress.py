import os

import pandas as pd
from xlwings import Workbook, Range
from win32com.client import Dispatch
import reporting


def compress_data():

    Workbook.caller()
    path = Range('Tools', 'ZZ1').value
    output_path = path[:path.rindex('\\')]

    data = pd.read_excel(path, 'data', index_col = None)
    sheet_name = 'Campaigns_Pivot_' + str(data['Date'].max()) + '.xlsm'
    joinpath = os.path.join(output_path, sheet_name)

    columns = list(data.columns)
    columns = columns[:columns.index('Post-Impression Activity') + 1]

    Range('Tools', 'ZZ2').value = joinpath
    xl = Dispatch('Excel.Application')
    xlwb = xl.Workbooks.Add()
    xlws = xlwb.Worksheets('Sheet1')
    xlws.Name = 'data'

    Range('Tools', 'ZZ1').horizontal.clear()

    xlOpenXMLWorkbookMacroEnabled = 52
    xlwb.SaveAs(joinpath, FileFormat = xlOpenXMLWorkbookMacroEnabled)

    wb = Workbook(joinpath)
    wb.set_current()
    data = data[columns]
    reporting.chunk_df(data, 'data', 'A1', 2000)






