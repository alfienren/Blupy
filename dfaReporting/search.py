from xlwings import Range, Sheet, Workbook
import pandas as pd
from reporting import datafunc, custom_variables


def query_and_cfv_data(path):
    query_sheets = ['Weekly Dash', 'Sitelink DDR', 'Sitelink Remarketing', 'Retention Location Intent Query',
                    'Location Intent Query']

    search_data = pd.DataFrame()

    for i in query_sheets:
        data = pd.read_excel(path, i, index_col=None)
        search_data = search_data.append(data)

    cfv = datafunc.read_cfv_report(path)
    cfv = custom_variables.custom_variable_columns(cfv)
    cfv = custom_variables.ddr_custom_variables(cfv)

    search_data = search_data.append(cfv)

    return search_data
