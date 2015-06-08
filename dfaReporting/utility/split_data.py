import pandas as pd
import os
from win32com.client import Dispatch
from xlwings import Workbook, Range
from weekly import data_output

def split():

    Workbook.caller()

    pivot_script = r'C:\Users\aarschle1\Google Drive\Optimedia\T-Mobile\Projects\Weekly_Reporting\bas\Pivot_Generate.bas'

    sheet = Range('Lookup', 'AA1').value
    column = Range('Lookup', 'AB1').value

    sheet_path = sheet[:sheet.rindex('\\')]
    folder_name = 'split data'
    path = os.path.join(sheet_path, folder_name)

    try:
        os.makedirs(path)
    except OSError:
        if not os.path.isdir(path):
            raise

    data = pd.read_excel(sheet, 'data', index_cols= None, na_values= [0])
    data.sort(column, axis= 0, inplace= True)

    split = data.groupby(column)

    xl = Dispatch('Excel.Application')
    xl.Visible = True

    xlOpenXMLWorkbookMacroEnabled = 52

    for i in range(0, len(split.groups)):

        xlwb = xl.Workbooks.Add()
        xlws = xlwb.Worksheets('Sheet1')
        xlws.Name = 'data'

        save_path = os.path.join(path, split.groups.keys()[i])
        xlwb.SaveAs(save_path, FileFormat = xlOpenXMLWorkbookMacroEnabled)

        wb = Workbook(save_path)
        wb.set_current()
        data_output.chunk_df(split.get_group(i), 'data', 'A1', 2500)
        xlwb.VBProject.VBComponents.Add(pivot_script)
        wb.save()
        wb.close()

