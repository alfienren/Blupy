import pandas as pd
import os
from win32com.client import Dispatch
from xlwings import Workbook, Range, Sheet
from weekly import data_output

def split():

    Workbook.caller()

    pivot_script = 'C:/Users/aarschle1/Google Drive/Optimedia/T-Mobile/Projects/Weekly_Reporting/bas/Tools_PivotGenerate.bas'

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
    j = 0

    for i in split.groups:

        xlwb = xl.Workbooks.Add()
        xlws = xlwb.Worksheets('Sheet1')
        xlws.Name = 'data'

        module = xlwb.VBProject.VBComponents.Add(1)
        module.CodeModule.AddFromFile(pivot_script)

        #wb = Workbook(xlwb.FullName)
        #wb.set_current()
        #data_output.chunk_df(split.get_group(i), 'data', 'A1', 2500)
        save_path = os.path.join(path, split.groups.keys()[j])
        data.to_excel(xlwb.FullName, 'data')
        xlwb.Application.Run(str(xlwb.Name) + '!Tools_PivotGenerate.GeneratePivot')
        xlwb.SaveAs(save_path, FileFormat = xlOpenXMLWorkbookMacroEnabled)

        j += 1
        xlwb.Close()
        #wb.save()
        #wb.close()