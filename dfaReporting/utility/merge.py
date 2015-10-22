import os
import ctypes

from xlwings import Workbook, Range, Sheet
import pandas as pd


def merge_data():

    Workbook.caller()

    if Range('data', 'A1').value is None:

        ctypes.windll.user32.MessageBoxA(0, 'data tab is empty', 'Data Missing in data tab', 0)
        exit()

    path = os.path.normpath(Range('Lookup', 'AA2').value)
    data_workbook = os.path.normpath(Range('Lookup', 'AB2').value)

    new_data = pd.read_excel(path, 'data', index_cols= None)

    if data_workbook[:-4] != '.csv':

        workbook_sheet = Range('Lookup', 'AC2').value
        data_to_merge = pd.read_excel(data_workbook, workbook_sheet, index_cols= None)

    else:

        data_to_merge = pd.read_csv(data_workbook, na_values=[0])

    data = data_to_merge.append(new_data)

    if Range('data', 'A1').value is None:

        Sheet('data').clear_contents()

    data_output.chunk_df(data, 'data', 'A1', 2500)


