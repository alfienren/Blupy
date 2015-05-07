__author__ = 'aarschle1'

from weekly import data_load, ExcelImport, cfv, categorization, action_tags, clickthroughs, floodlight, dfa_data_output
from xlwings import Workbook
import os

def macro():

    wb = Workbook.caller()

    sa = data_load.sa_data()

    floodlights = data_load.cfv_data()
    floodlights = cfv.cfv_munge(floodlights)

    data = sa.append(floodlights)
    data = clickthroughs.clickthrough(data)
    data = floodlight.floodlight(data)
    data = action_tags.actions(data)
    data = categorization.categories(data)

    ExcelImport.chunk_df(data, 'working', 'A1', 5000)


    wb = Workbook()

    ExcelImport.chunk_df(data, 'Sheet1', 'A1', 5000)

if __name__ == '__main__':
    # To run from Python, not needed when called from Excel.
    # Expects the Excel file next to this source file, adjust accordingly.
    path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'dfa_test.xlsm'))
    Workbook.set_mock_caller(path)
    macro()
