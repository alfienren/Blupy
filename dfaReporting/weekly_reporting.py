from weekly import data_load, cfv_clean, categorization, action_tags, clickthroughs, floodlight_transform, \
    data_output, f_tags, messages
from xlwings import Workbook, Range, Sheet
import pandas as pd

def weekly_reporting():

    wb = Workbook.caller()
    wb.save()
    sa = data_load.sa()

    cfv_variables = data_load.cfv()
    cfv_variables = cfv_clean.cfv_data(cfv_variables)

    data = sa.append(cfv_variables)
    data = clickthroughs.clickthrough(data)
    data = floodlight_transform.floodlight_data(data)
    data = action_tags.actions(data)
    data = categorization.categories(data)
    data = categorization.sites(data)
    data = messages.messaging(data)
    data = f_tags.f_tags(data)
    data = data_output.output(data)

    columns = data_output.columns(sa)
    data = data[columns]
    data.fillna(0, inplace=True)

    if Range('data', 'A1').value is None:
        data_output.chunk_df(data, 'data', 'A1', 5000)

    # If data is already present in the tab, the two data sets are merged together and then copied into the data tab.
    else:
        past_data = pd.DataFrame(pd.read_excel(Range('Lookup', 'AA1').value, 'data', index_col=None))
        appended_data = past_data.append(data)
        appended_data = appended_data[columns]
        appended_data.fillna(0, inplace=True)
        Sheet('data').clear()
        data_output.chunk_df(appended_data, 'data', 'A1', 5000)

'''
if __name__ == '__main__':
    # To run from Python, not needed when called from Excel.
    # Expects the Excel file next to this source file, adjust accordingly.
    path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'DFA_Weekly_Reporting.xlsm'))
    Workbook.set_mock_caller(path)
    weekly_reporting()
'''