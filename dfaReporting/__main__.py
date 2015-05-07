from weekly import data_load, ExcelImport, cfv, categorization, action_tags, clickthroughs, floodlight, \
    dfa_data_output, f_tags
from xlwings import Workbook, Range, Sheet
import os
import pandas as pd

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
    data = f_tags.f_tags(data)
    data = dfa_data_output.output(data)

    columns = data_load.columns(sa)
    data = data[columns]
    data.fillna(0, inplace=True)

    if Range('data', 'A1').value is None:
        ExcelImport.chunk_df(data, 'data', 'A1', 5000)

    # If data is already present in the tab, the two data sets are merged together and then copied into the data tab.
    else:
        past_data = pd.DataFrame(pd.read_excel(Range('Lookup', 'AA1').value, 'data', index_col=None))
        appended_data = past_data.append(data)
        appended_data = appended_data[columns]
        appended_data.fillna(0, inplace=True)
        Sheet('data').clear()
        ExcelImport.chunk_df(appended_data, 'data', 'A1', 5000)

if __name__ == '__main__':
    # To run from Python, not needed when called from Excel.
    # Expects the Excel file next to this source file, adjust accordingly.
    path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'dfa_test.xlsm'))
    Workbook.set_mock_caller(path)
    macro()
