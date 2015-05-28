import pandas as pd
from xlwings import Workbook, Range, Sheet
from weekly import data_output
import os
from win32com.client import Dispatch

def compress_data():

    Workbook.caller()
    path = Range('Tools', 'ZZ1').value
    output_path = path[:path.rindex('\\')]
    joinpath = os.path.join(output_path, 'compressed_data.xlsm')
    data = pd.read_excel(path, 'data', index_col = None)

    columns = list(data.columns)
    columns = columns[:columns.index('Post-Impression Activity') + 1]

    Range('Tools', 'ZZ2').value = joinpath
    xl = Dispatch('Excel.Application')
    xlwb = xl.Workbooks.Add()
    xlws = xlwb.Worksheets('Sheet1')
    xlws.Name = 'data'
    xlOpenXMLWorkbookMacroEnabled = 52
    xlwb.SaveAs(joinpath, FileFormat = xlOpenXMLWorkbookMacroEnabled)
    xlwb.Close()

    wb = Workbook(joinpath)
    wb.set_current()
    data = data[columns]
    data_output.chunk_df(data, 'data', 'A1', 2000)




