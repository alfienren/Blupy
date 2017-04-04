import os
import re

import numpy as np
import pandas as pd
from xlwings import Range, Workbook, Sheet

from categorization import Categorization


class DataMethods(object):

    def __init__(self):
        self.wb = Workbook.caller()
        self.path = self.wb.fullname
        self.temp_site_activity = 'SA_Temp'
        self.temp_cfv = 'CFV_Temp'
        self.action_reference = 'Action_Reference'
        self.lookup = 'Lookup'

    def read_site_activity_report(self, adv='tmo'):
        self.sa = pd.read_excel(self.path, self.temp_site_activity, index_col=None)
        if 'DBM Cost USD' in list(self.sa.columns):
            self.sa.rename(columns={'DBM Cost USD': 'DBM Cost (USD)'}, inplace=True)

        if adv == 'tmo':
            sa_creative = self.sa[['Placement', 'Creative Field 1']]
            sa_creative = sa_creative.drop_duplicates(subset='Placement')

            return self.sa, sa_creative

        else:
            return self.sa

    def read_cfv_report(self):
        self.cfv = pd.read_excel(self.path, self.temp_cfv, index_col=None)

        return self.cfv

    def generate_search_cfv_report(self):
        cfv = pd.read_excel(self.path, 'RAW', index_col=None)

        device_lookup = pd.read_table(self.path[:self.path.rindex('\\')] + '\\devices.txt')
        device_lookup.set_index('Device SKU', inplace=True)

        device_string = cfv['Device (string)'].str.split(',').apply(pd.Series).stack()
        device_string.index = device_string.index.droplevel(-1)
        device_string.name = "Device ID"

        cfv = cfv[['Paid Search Engine Account', 'ORD Value', 'Date', 'Device (string)']]
        cfv['Paid Search Engine Account'] = np.where(pd.isnull(cfv['Paid Search Engine Account']) == True, 'Whistleout',
                                                     cfv['Paid Search Engine Account'])

        undefined = cfv[cfv['ORD Value'] == 'undefined']
        cfv = cfv[(cfv['ORD Value'] != 'undefined') & (pd.isnull(cfv['Device (string)']) != True)]
        del cfv['Device (string)']

        device_cfv = cfv.join(device_string)

        device_cfv['Device ID'].fillna(0, inplace=True)
        device_cfv.set_index('Device ID', inplace=True)

        cfv_new = pd.merge(device_cfv, device_lookup, how='left', left_index=True, right_index=True).reset_index()
        #cfv_new = cfv_new[(cfv_new['Product Subcategory'] == 'Postpaid') |
        # (cfv_new['Product Subcategory'] == 'Prepaid')]

        cfv_new = cfv_new.append(undefined)
        cfv_new = Categorization().date_columns(cfv_new)
        del cfv_new['Quarter']

        cols = {'Paid Search Engine Account': 'Account',
                'ORD Value': 'Order Number',
                'Product Name': 'Title',
                'Manufacturer': 'Brand',
                'Product Category': 'Product Type',
                'Product Subcategory': 'Postpaid/Prepaid',
                'index': 'Device ID'}

        cfv_new.rename(columns=cols, inplace=True)

        cfv_new = cfv_new[['Month', 'Week', 'Date', 'Account', 'Order Number', 'Device ID',
                           'Title', 'Brand', 'Product Type', 'Postpaid/Prepaid']]

        Sheet('output').clear_contents()

        DataMethods().chunk_df(cfv_new, 'output', 'A1')

    @staticmethod
    def chunk_df(df, sheet, startcell, chunk_size=5000):
        if len(df) <= (chunk_size + 1):
            Range(sheet, startcell, index=False, header=True).value = df

        else:
            Range(sheet, startcell, index=False).value = list(df.columns)
            c = re.match(r"([a-z]+)([0-9]+)", startcell[0] + str(int(startcell[1]) + 1), re.I)
            row = c.group(1)
            col = int(c.group(2))

            for chunk in (df[rw:rw + chunk_size] for rw in
                          range(0, len(df), chunk_size)):
                Range(sheet, row + str(col), index=False, header=False).value = chunk
                col += chunk_size

    @staticmethod
    def save_search_raw_data(search_pivoted):
        save_path = Range('Lookup', 'K1').value
        through_date = search_pivoted['Date'].max().strftime('%m.%d.%Y')
        file_name = 'Search_Raw_Data_' + through_date + '.xlsx'
        search_pivoted.to_excel(os.path.join(save_path, file_name), index=False)

    @staticmethod
    def merge_past_data(data, columns, path):
        if Range('data', 'A1').value is None:
            DataMethods().chunk_df(data, 'data', 'A1')

        # If data is already present in the tab, the two data sets are merged together and then copied into the data tab.

        else:
            past_data = pd.read_excel(path, 'data', index_col=None)
            appended_data = past_data.append(data)
            appended_data = appended_data[columns]
            appended_data.fillna(0, inplace=True)
            Sheet('data').clear_contents()
            DataMethods().chunk_df(appended_data, 'data', 'A1')

