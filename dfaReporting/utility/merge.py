from weekly import data_output
from xlwings import Workbook, Range, Sheet
import pandas as pd
import os
import ctypes

def merge_data():

    Workbook.caller()

    if Range('data', 'A1').value is None:

        ctypes.windll.user32.MessageBoxA(0, 'data tab is empty', 'Data Missing in data tab', 0)
        exit()

    path = os.path.normpath(Range('Lookup', 'AA2').value)
    data_workbook = os.path.normpath(Range('Lookup', 'AB2').value)
    workbook_sheet = Range('Lookup', 'AC2').value

    new_data = pd.read_excel(path, 'data', index_cols= None)
    data_to_merge = pd.read_excel(data_workbook, workbook_sheet, index_cols= None)

    data = data_to_merge.append(new_data)

    if Range('data', 'A1').value is None:

        data_output.chunk_df(data, 'data', 'A1', 2500)

    else:

        Sheet('data').clear_contents()
        data_output.chunk_df(data, 'data', 'A1', 2500)


