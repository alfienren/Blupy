from xlwings import Range, Sheet, Workbook
import pandas as pd
import numpy as np
import os
import sys
from reporting import datafunc, custom_variables, categorization


def query_and_cfv_data(path):
    query_sheets = ['Weekly Dash', 'Sitelink DDR', 'Sitelink Remarketing', 'Sitelink Bing',
                    'Retention Location Intent Query', 'Location Intent Query']

    search_data = pd.DataFrame()

    for i in query_sheets:
        data = pd.read_excel(path, i, index_col=None)

        if i == 'Sitelink Bing':

            dr_brand = '|'.join(list(['DDR_B_High_Volume', 'DDR_B_Bring']))
            remarket = '|'.join(list(['REM_']))
            b_marketing = '|'.join(list(['DDR_B_']))
            data.rename(columns={'Campaign name':'Campaign', 'Ad extension property value':'Sitelink display text'},
                        inplace=True)

            data = data[(data['Sitelink display text'].str.contains('My T-Mobile') == True)
                        | (data['Sitelink display text'].str.contains('for Business') == True)]

            data['Bucket Class'] =  np.where(data['Campaign'].str.contains(dr_brand) == True, 'DR-Brand',
                                             np.where(data['Campaign'].str.contains(remarket) == True, 'Remarketing',
                                                      np.where(data['Campaign'].str.contains(b_marketing) == True,
                                                               'Brand Marketing', None)))

        data['Source'] = i
        search_data = search_data.append(data)

    search_data.rename(columns={'From':'Date'}, inplace=True)

    search_data['Week'] = search_data['Date'].apply(lambda x: categorization.mondays(x))

    cfv = datafunc.read_cfv_report(path)
    cfv['Source'] = 'CFV'

    cfv = custom_variables.custom_variable_columns(cfv)
    cfv = custom_variables.ddr_custom_variables(cfv)

    search_data = search_data.append(cfv)

    search_data.fillna(0, inplace=True)

    search_pivoted = pd.pivot_table(search_data, index=['Source', 'Account', 'Bucket Class', 'Business Unit', 'Campaign',
                                          'Device segment', 'Date', 'Sitelink display text', 'Accessory (string)',
                                          'OrderNumber (string)', 'Device (string)', 'Plan (string)',
                                          'Service (string)'],
                             aggfunc=np.sum, fill_value=0).reset_index()

    search_pivoted['Location Intent_temp'] = np.where(search_pivoted['Source'] == 'Weekly Dash',
                                                      search_pivoted['Store Locator'], 0)
    search_pivoted['Location Intent_temp2'] = np.where(search_pivoted['Source'] == 'Location Intent Query',
                                                       search_pivoted['Clicks'], 0)

    search_pivoted['Location Intent'] = search_pivoted['Location Intent_temp'] + search_pivoted['Location Intent_temp2']
    search_pivoted.drop(['Location Intent_temp', 'Location Intent_temp2'], axis=1, inplace=True)

    search_pivoted.rename(columns={'Cost':'Spend', 'Impr':'Impressions'}, inplace=True)

    search_pivoted = categorization.search_lookup(search_pivoted)
    search_pivoted = categorization.date_columns(search_pivoted)

    datafunc.chunk_df(search_pivoted, 'data', 'A1')


def save_raw_data_file(search_pivoted):
    save_path = Range('Lookup', 'K1').value

    through_date = search_pivoted['Date'].max().strftime('%m.%d.%Y')

    file_name = 'Search_Raw_Data_' + through_date + '.xlsx'

    search_pivoted.to_excel(os.path.join(save_path, file_name), index=False)


def generate_search_reporting():
    wb = Workbook.caller()
    wb.save()

    query_and_cfv_data(wb.fullname)




