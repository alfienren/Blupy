import paths
import numpy as np
import pandas as pd
import datetime
import re
from xlwings import Range, Sheet


def raw_pivot():
    ddr = pd.read_excel(paths.dr_pivot_path(), 'Working Data', index_cols=None, parse_cols='A:V,X,Z:AK,CP:DH')
    ddr.fillna(0, inplace=True)

    return ddr


def tableau_campaign_data(ddr):
    ddr_columns = ['Campaign', 'Week', 'Site', 'Message Tactic', 'Placement Messaging Type', 'Message Offer', 'A', 'B',
                   'C', 'D', 'SLV', 'Awareness Actions', 'Consideration Actions', 'Total Traffic Actions',
                   'NTC Media Cost', 'NET Media Cost', 'Impressions', 'Clicks', 'Orders', 'Total GAs', 'Prepaid GAs',
                   'Postpaid GAs', 'Prepaid SIMs', 'Postpaid SIMs', 'Prepaid Mobile Internet',
                   'Postpaid Mobile Internet', 'Prepaid phone', 'Postpaid phone', 'AAL', 'New device']

    ddr.rename(columns={'Tactic': 'Message Tactic'}, inplace=True)

    ddr['Week'] = pd.to_datetime(ddr['Week'])
    end = ddr['Week'].max()
    delta = datetime.timedelta(weeks=0)
    start = end - delta
    ddr_data = ddr[(ddr['Week'] >= start) & (ddr['Week'] <= end)]

    ddr_data = ddr_data[ddr_columns]

    gb = ddr_data.groupby(['Campaign', 'Week', 'Site', 'Message Tactic', 'Placement Messaging Type', 'Message Offer'])

    grouped = gb.aggregate(np.sum).reset_index()

    grouped['Tactic'] = 'Display'
    grouped['Channel'] = np.where(grouped['Campaign'] == 'DR', 'DR', 'Brand Remessaging')

    grouped['Source'] = 'DR-Pivot'

    return grouped


def tableau_pacing(forecast_data):
    column_names = '|'.join(list(['ASG', 'Amazon', 'Magnetic', 'eBay', 'Yahoo!', 'Bazaar Voice', 'Date', 'Type']))
    sites = '|'.join(list(['ASG', 'Amazon', 'Bazaar Voice', 'eBay', 'Magnetic', 'Yahoo!']))
    metrics = '|'.join(list(['Spend', 'GAs']))
    to_remove = '|'.join(list([sites, metrics]))
    conf_interval = '|'.join(list(['Hi', 'Lo']))

    tableau_data = forecast_data.select(lambda x: re.search(column_names, x), axis= 1)

    tableau_data = pd.melt(tableau_data, id_vars= ['Date', 'Type'],
                   value_vars= [col for col in tableau_data.columns if re.search(sites, col)])

    tableau_data['Site'] = tableau_data['variable'].str.split(' ').str[0]
    tableau_data['Metric'] = tableau_data['variable'].str.split(' ').str[-1]
    tableau_data['variable'] = tableau_data['variable'].str.replace(to_remove, '').str.strip()

    tableau_data.rename(columns={'variable':'Tactic'}, inplace= True)
    tableau_data = tableau_data[tableau_data['Tactic'].str.contains(conf_interval) == False]

    tableau_data_output = pd.pivot_table(tableau_data, index=['Site', 'Tactic', 'Date', 'Type'],
              columns=['Metric'], values='value', aggfunc=np.sum)
    tableau_data_output.reset_index(inplace= True)

    Sheet('tableau_pacing_data').clear_contents()

    Range('tableau_pacing_data', 'A1', index= False).value = tableau_data_output

    return tableau_data


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

    reporting.chunk_df(client_data, 0, 'A1')

    wb.save(save_path + '\\' + 'DR_Search_Raw_Data.csv')
    wb.close()