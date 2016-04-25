import numpy as np
import pandas as pd


def search_data_client(search_data, save_path):
    column_order = ['Search Engine', 'Brand DDR Bucket', 'Week', 'NET Media Cost', 'Impressions', 'Clicks', 'Orders',
                    'Plans', 'Prepaid Orders', 'Consideration Actions', 'Add-A-Line', 'Total GAs', 'New Total eGAs',
                    'Telesales GAs']

    client_data = search_data[column_order]
    client_data = client_data[client_data['Week'] >= '1/1/2016']

    client_data['Brand DDR Bucket']  = np.where(pd.isnull(client_data['Brand DDR Bucket']) == True,
                                                client_data['Search Engine'],
                                                client_data['Brand DDR Bucket'])

    client_data.to_csv(save_path + '\\' + 'DR_Search_Raw_Data.txt', sep='\t', encoding='utf-8')


def display_data_client(dr_data, save_path):
    dr_data.to_csv(save_path + '\\' + 'DDR_Raw_Data.txt', sep='\t', encoding='utf-8')
