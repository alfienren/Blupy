import os
import re

import numpy as np
import pandas as pd
from xlwings import Range, Workbook, Sheet

from analytics.data.categorization import Categorization


class DataMethods(object):

    def __init__(self):
        super(DataMethods, self).__init__()
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

    @staticmethod
    def additional_columns(data, adv='tmo'):
        # The DFA field DBM Cost is more accurate for placements using dynamic bidding. If a placement is not using
        # dynamic bidding, DBM Cost = 0. Therefore, if DBM cost does not equal 0, replace the row's media cost with
        # DBM cost. If DBM Cost = 0, Media Cost stays the same.

        if adv == 'tmo' or adv == 'dr':
            data['Media Cost'] = np.where(data['DBM Cost (USD)'] != 0, data['DBM Cost (USD)'], data['Media Cost'])
        if adv == 'tmo':
            data.drop('DBM Cost (USD)', 1, inplace=True)
        if adv == 'dr':
            data.rename(columns={'Campaign': 'Campaign2'}, inplace=True)
            data['Campaign'] = np.where(data['Campaign2'].str.contains('DDR') == True, 'DR', 'Brand Remessaging')
            data['NET Media Cost'] = data['Media Cost']

        if adv != 'dr':
            data['Video Completions'] = 0
            data['Video Views'] = 0

        data['NTC Media Cost'] = 0

        return data

    @staticmethod
    def order_columns(adv='tmo'):
        if adv == 'tmo':
            dimensions = ['Week', 'Date', 'Month', 'Quarter', 'Campaign', 'Media Plan', 'Language', 'Site (DCM)',
                          'Site',
                          'Click-through URL', 'F Tag', 'Category', 'Category_Adjusted', 'Message Bucket',
                          'Message Category', 'Creative Bucket', 'Creative Theme', 'Creative Type',
                          'Creative Groups 1',
                          'Creative ID', 'Ad', 'Creative Groups 2', 'Message Campaign', 'Creative Field 1',
                          'Placement Messaging Type', 'Placement', 'Placement ID', 'Placement Cost Structure']

            cfv_floodlight_columns = ['OrderNumber (string)', 'Activity', 'Floodlight Attribution Type',
                                      'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)']

            metrics = ['Media Cost', 'NTC Media Cost', 'Impressions', 'Clicks', 'Orders', 'Plans', 'Add-a-Line',
                       'Activations', 'Devices', 'Services', 'Accessories', 'Postpaid Plans', 'Prepaid Plans',
                       'eGAs',
                       'Store Locator Visits', 'A Actions', 'B Actions', 'C Actions', 'D Actions', 'E Actions',
                       'F Actions',
                       'Awareness Actions', 'Consideration Actions', 'Traffic Actions', 'Post-Click Activity',
                       'Post-Impression Activity', 'Video Views', 'Video Completions', 'Prepaid GAs',
                       'Postpaid GAs',
                       'Postpaid Orders', 'Prepaid Orders', 'Prepaid SIMs', 'Postpaid SIMs',
                       'Prepaid Mobile Internet',
                       'Postpaid Mobile Internet', 'Prepaid Phone', 'Postpaid Phone', 'Total GAs',
                       'DDR New Devices',
                       'DDR Add-a-Line']

            new_columns = dimensions + metrics + cfv_floodlight_columns

        else:
            dimensions = ['Week', 'Date', 'Month', 'Quarter', 'Campaign', 'Language', 'Site (DCM)', 'Site',
                          'TMO_Category', 'TMO_Category_Adjusted', 'Creative', 'Creative Type', 'Creative Groups 1',
                          'Creative ID', 'Ad', 'Creative Groups 2', 'Creative Field 2', 'Placement', 'Placement ID',
                          'Category', 'Creative Type Lookup', 'Skippable']

            cfv_floodlight_columns = ['Floodlight Attribution Type', 'Activity', 'Transaction Count']

            metrics = ['Media Cost', 'NTC Media Cost', 'Impressions', 'Clicks', 'Orders', 'Store Locator Visits',
                       'GM A Actions', 'GM B Actions', 'GM C Actions', 'GM D Actions', 'Hispanic A Actions',
                       'Hispanic B Actions', 'Hispanic C Actions', 'Hispanic D Actions', 'Total A Actions',
                       'Total B Actions', 'Total C Actions', 'Total D Actions', 'Awareness Actions',
                       'Traffic Actions',
                       'Consideration Actions', 'Post-Click Activity', 'Post-Impression Activity', 'Video Views',
                       'Video Completions']

            new_columns = dimensions + metrics + cfv_floodlight_columns

        if adv == 'dr':
            dimensions1 = ['Campaign', 'Month', 'Week', 'Site', 'Tactic', 'Category', 'Placement Messaging Type',
                           'Message Bucket', 'Message Category', 'Message Offer']

            dimensions2 = ['Campaign2', 'Date', 'Site (DCM)', 'Creative', 'Click-through URL',
                           'Creative Pixel Size',
                           'Creative Type', 'Creative Field 1', 'Ad', 'Creative Groups 2', 'Placement',
                           'Placement ID',
                           'Placement Cost Structure']

            metrics1 = ['A Actions', 'B Actions', 'C Actions', 'D Actions', 'Store Locator Visits',
                        'Awareness Actions',
                        'Consideration Actions', 'Traffic Actions', 'Post-Impression Activity',
                        'Post-Click Activity',
                        'NTC Media Cost', 'NET Media Cost']

            metrics2 = ['Impressions', 'Clicks', 'Media Cost', 'DBM Cost (USD)']

            new_columns = dimensions1 + metrics1 + dimensions2 + metrics2

        return list(new_columns)

    @staticmethod
    def dr_drop_columns(dr):
        cols_to_drop = ['Month', 'Tactic', 'Category', 'Message Bucket', 'Message Category',
                        'Message Offer', 'A Actions', 'B Actions', 'C Actions', 'D Actions', 'Store Locator Visits',
                        'Awareness Actions', 'Consideration Actions', 'Post-Impression Activity',
                        'Post-Click Activity',
                        'NET Media Cost', 'Clicks', 'Prepaid GAs', 'Postpaid GAs', 'Prepaid SIMs', 'Postpaid SIMs',
                        'Prepaid Mobile Internet', 'Postpaid Mobile Internet', 'Prepaid Phone', 'Postpaid Phone',
                        'DDR Add-a-Line', 'DDR New Devices']

        ddr = dr.drop(cols_to_drop, axis=1)

        return ddr

    @staticmethod
    def strip_clickthroughs(data):

        data['Click-through URL'] = data['Click-through URL'].str.replace('http://analytics.bluekai.com/site/', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('%3F%3DADV_DS_%epid!_%eaid!_%ecid!_%eadv!',
                                                                          '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('15991\?phint', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('http://15991\?phint', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('event%3Dclick&phint', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('aid%3D%eadv!&phint', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('pid%3D%epid!&phint', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('cid%3D%ebuy!&phint', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('crid%3D%ecid!&done', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('pid%3D%25epid!&phint', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('%3D%epid!_%eaid!_%ecid!_%eadv!', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('%26csdids', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('DADV_DS_ADDDVL4Q_EMUL7Y9E1YA4116', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('DWTR_DD_DDRDSPLYPR_JM2694TSP3U5895', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('DWTR_DD_DDRFCBK_RQLMKXRCUQZ1042', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('%3Fcmpid%3', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('b/refmh_', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace(
            '%3Fcm_mmc%3DPurchase-_-Display-_-Revere-_-Revere', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace(
            '%3Fcm_mmc%3DDisplay-_-Purchase-_-GM-_-Tablet_Base', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace(
            '%3Fcmpid%3DWTR_DD_DDRFCBK_RQLMKXRCUQZ1042%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace(
            '%3F%26csdids%3DADV_DS_%epid!_%eaid!_%ecid!_%eadv!', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace(
            '%3Fcm_mmc%3DPurchase-_-Display-_-Revere-_-Revere%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace(
            '%3Fcm_mmc%3DDisplay-_-Purchase-_-GM-_-Tablet_Base%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace(
            '%3Fcmpid%3DWTR_DD_DDRDSPLYPR_JM2694TSP3U5895%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('&csdids%epid!_%eaid!_%ecid!_%eadv!', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('=', '')
        data['Click-through URL'] = data['Click-through URL'].str.replace('%2F', '/')
        data['Click-through URL'] = data['Click-through URL'].str.replace('%3A', ':')
        data['Click-through URL'] = data['Click-through URL'].str.replace('%23', '#')
        data['Click-through URL'] = data['Click-through URL'].apply(lambda x: str(x).split('.html')[0])
        data['Click-through URL'] = data['Click-through URL'].apply(lambda x: str(x).split('?')[0])
        data['Click-through URL'] = data['Click-through URL'].apply(lambda x: str(x).split('%')[0])
        data['Click-through URL'] = data['Click-through URL'].apply(lambda x: str(x).split('_')[0])
        data['Click-through URL'] = data['Click-through URL'].str.replace('DWTR', '')

        return data
