import pandas as pd
import numpy as np
from xlwings import Workbook, Range
import main


def merge_search_data():
    sheet = Range('Sheet3', 'AC1').value

    search_tabs = ['Search Raw Data', 'Whistleout', 'CSE', 'Ad Marketplace']

    search_data = pd.DataFrame()

    for i in search_tabs:
        df = pd.read_excel(sheet, i, index_cols=None)
        df['Source'] = i
        search_data = search_data.append(df)

    return search_data


def tableau_search_data(search_data):
    search_data[['Total GAs', 'New Total eGAs']] = search_data[['Total GAs', 'New Total eGAs']].fillna(0)

    search_data['Total GAs'] = search_data['Total GAs'] + search_data['New Total eGAs']
    search_data['NTC Media Cost'] = search_data['NET Media Cost'] / .96759

    search_gas = pd.DataFrame(Range('Search_GAs', 'A1').table.value,
                              columns= Range('Search_GAs', 'A1').horizontal.value)
    search_gas.drop(0, inplace= True)

    search_gas = search_gas[['Week', 'Total Traffic Actions', 'Prepaid GAs', 'Postpaid GAs', 'Prepaid SIMs',
                             'Postpaid SIMs']]

    search_gas['Source'] = 'Search Raw Data'

    search = search_data.append(search_gas)

    search['Tactic'] = 'Search'
    search['Channel'] = 'DR'
    search['Campaign'] = 'Search'

    return search


def search_data_client(search_data, save_path):
    column_order = ['Search Engine', 'Brand DDR Bucket', 'Week', 'NET Media Cost', 'Impressions', 'Clicks', 'Orders',
                    'Plans', 'Prepaid Orders', 'Consideration Actions', 'Add-A-Line', 'Total GAs', 'New Total eGAs',
                    'Telesales GAs']

    client_data = search_data[column_order]
    client_data = client_data[client_data['Week'] >= '1/1/2016']

    client_data['Brand DDR Bucket']  = np.where(pd.isnull(client_data['Brand DDR Bucket']) == True,
                                                client_data['Search Engine'],
                                                client_data['Brand DDR Bucket'])

    wb = Workbook()
    wb.set_current()

    main.chunk_df(client_data, 0, 'A1')

    wb.save(save_path + '\\' + 'DR_Search_Raw_Data.csv')
    wb.close()
