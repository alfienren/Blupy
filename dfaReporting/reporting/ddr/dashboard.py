import datetime
import re

import numpy as np
import pandas as pd
from xlwings import Range, Sheet, Workbook

from reporting import datafunc
from reporting.ddr import client_raw_data


def load_raw_dr_data():
    path = Range('Sheet3', 'AB1').value

    ddr = pd.read_excel(path, 'data', index_cols=None, parse_cols='A:V,X,Z:AK,CR:DL')
    ddr.fillna(0, inplace=True)

    return [ddr, path]


def dr_display_data(ddr):
    ddr_columns = ['Campaign', 'Week', 'Site', 'Message Tactic', 'Placement Messaging Type', #'Message Offer',
                   'A Actions', 'B Actions', 'C Actions', 'D Actions', 'Store Locator Visits', 'Awareness Actions',
                   'Consideration Actions', 'Traffic Actions', 'View-through Conversions', 'Click-through Conversions',
                   'NTC Media Cost', 'NET Media Cost', 'Impressions', 'Clicks', 'Orders', 'Prepaid Orders',
                   'Postpaid Orders', 'Total GAs', 'Prepaid GAs', 'Postpaid GAs', 'Prepaid SIMs', 'Postpaid SIMs',
                   'Prepaid Mobile Internet', 'Postpaid Mobile Internet', 'Prepaid Phone', 'Postpaid Phone',
                   'DDR Add-a-Line', 'DDR New Devices']

    ddr.rename(columns={'Tactic': 'Message Tactic',
                        'Post-Impression Activity': 'View-through Conversions',
                        'Post-Click Activity' : 'Click-through Conversions'}, inplace=True)

    # ddr['Week'] = pd.to_datetime(ddr['Week'])
    # end = ddr['Week'].max()
    # delta = datetime.timedelta(weeks=0)
    # start = end - delta
    # ddr_data = ddr[(ddr['Week'] >= start) & (ddr['Week'] <= end)]

    ddr_data = ddr[ddr_columns]

    telesales = pd.DataFrame(Range('Telesales', 'A1').table.value, columns=Range('Telesales', 'A1').horizontal.value)
    telesales.drop(0, inplace=True)
    # telesales = telesales[(telesales['Week'] >= start) & (telesales['Week'] <= end)]
    telesales.set_index(['Site', 'Placement Messaging Type', 'Week'], inplace=True)

    gb = ddr_data.groupby(['Campaign', 'Week', 'Site', 'Message Tactic', 'Placement Messaging Type'])
    grouped = gb.aggregate(np.sum).reset_index()

    grouped.set_index(['Site', 'Placement Messaging Type', 'Week'], inplace=True)

    grouped = pd.merge(grouped, telesales, how='left', right_index=True, left_index=True)

    grouped.reset_index(inplace=True)

    grouped['Tactic'] = 'Display'
    grouped['Channel'] = np.where(grouped['Campaign'] == 'DR', 'DR', 'Brand Remessaging')

    grouped['Source'] = 'DR-Pivot'

    return grouped


def merge_search_data():
    sheet = Range('Sheet3', 'AC1').value

    search_tabs = ['Search Raw Data', 'Whistleout', 'CSE', 'Ad Marketplace']

    search_data = pd.DataFrame()

    for i in search_tabs:
        df = pd.read_excel(sheet, i, index_cols=None)
        df['Source'] = i
        search_data = search_data.append(df)

    return search_data


def dr_search_data(search_data):
    search_data[['Total GAs', 'New Total eGAs']] = search_data[['Total GAs', 'New Total eGAs']].fillna(0)

    search_data['Total GAs'] = search_data['Total GAs'] + search_data['New Total eGAs']
    search_data['NTC Media Cost'] = search_data['NET Media Cost'] / .96759

    search_gas = pd.DataFrame(Range('Search_GAs', 'A1').table.value,
                              columns= Range('Search_GAs', 'A1').horizontal.value)
    search_gas.drop(0, inplace= True)

    search_gas = search_gas[['Week', 'Traffic Actions', 'Prepaid GAs', 'Postpaid GAs', 'Prepaid SIMs',
                             'Postpaid SIMs']]

    search_gas['Source'] = 'Search Raw Data'

    search = search_data.append(search_gas)

    search['Tactic'] = 'Search'
    search['Channel'] = 'DR'
    search['Campaign'] = 'Search'

    return search


def generate_data():
    wb = Workbook.caller()

    dr_pivot = load_raw_dr_data()

    save_path = dr_pivot[1]
    save_path = save_path[:save_path.rindex('\\')]

    ddr_data = dr_pivot[0]

    ddr_display = dr_display_data(ddr_data)
    ddr_search_data = merge_search_data()

    tableau_search = dr_search_data(ddr_search_data)

    tableau = ddr_display.append(tableau_search)

    tableau['Quarter'] = None
    tableau['Week of Quarter'] = None

    if Range('merged', 'A1').value is None:
        datafunc.chunk_df(tableau, 'merged', 'A1')

    # If data is already present in the tab, the two data sets are merged together and then copied into the data tab.
    else:
        Sheet('merged').clear()
        datafunc.chunk_df(tableau, 'merged', 'A1')

    client_raw_data.search_data_client(ddr_search_data, save_path)
    client_raw_data.display_data_client(ddr_data, save_path)
