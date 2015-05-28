import pandas as pd
from xlwings import Workbook, Range, Sheet
from weekly import data_output
import os

def compress_data():

    path = Range('Tools', 'ZZ1').value
    workbook_path = os.path.normpath(path)
    output_path = os.path.normpath(path[:path.rindex('\\')])

    data = pd.read_excel(workbook_path, 'data', index_col = None)

    columns = list(data.columns)
    columns = columns[columns.index('Post-Impression Activity') + 1:]

    data = data[columns]

    wb = Workbook()

    Sheet(1).name = 'data'

    data_output.chunk_df(data, 'data', 'A1', 2000)

    wb.save(output_path + 'compressed_data.xlsx')
    wb.close()



