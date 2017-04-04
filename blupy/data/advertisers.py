import itertools
import re

import arrow
import numpy as np
import pandas as pd
from xlwings import Workbook, Range, Sheet

from categorization import Categorization
from file_io import DataMethods
from floodlights import Floodlights
from reporting.qa import QA


class UpdateAdvertisers(Categorization, Floodlights, DataMethods):

    def __init__(self):
        super(UpdateAdvertisers, self).__init__()

    def tmo(self):
        self.wb.save()

        # Before function is ran, VBA code will create the necessary tabs in order to process correctly. See the
        # documentation in the VBA modules for more information.
        # Workbook needs to be saved in order to load the data into pandas properly
        # Load the Site Activity and Custom Floodlight Variable data into pandas as DataFrames

        sa, sa_creative = DataMethods().read_site_activity_report(adv='tmo')

        cfv = pd.merge(DataMethods().read_cfv_report(), sa_creative, how='left', on='Placement')

        cfv = Floodlights().custom_variables(cfv)
        cfv = Floodlights().ddr_custom_variables(cfv)

        data = sa.append(cfv)
        data = self.strip_clickthroughs(data)

        data = Floodlights().a_e_traffic(data, adv='tmo')

        data = Categorization().categorize_report(data, adv='tmo')
        data = Floodlights().f_tags(data)
        data = self.additional_columns(data, adv='tmo')

        sa_columns = list(sa.columns)
        tag_columns = sa_columns[sa_columns.index('DBM Cost (USD)') + 1:]

        dimensions = ['Week', 'Date', 'Month', 'Quarter', 'Campaign', 'Media Plan', 'Language', 'Site (DCM)',
                      'Site', 'Click-through URL', 'F Tag', 'Category', 'Category_Adjusted', 'Message Bucket',
                      'Message Category', 'Creative Bucket', 'Creative Theme', 'Creative Type', 'Creative Groups 1',
                      'Creative ID', 'Ad', 'Creative Groups 2', 'Message Campaign', 'Creative Field 1',
                      'Placement Messaging Type', 'Placement', 'Placement ID', 'Placement Cost Structure']

        cfv_floodlight_columns = ['OrderNumber (string)', 'Activity', 'Floodlight Attribution Type',
                                  'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)']

        metrics = ['Media Cost', 'NTC Media Cost', 'Impressions', 'Clicks', 'Orders', 'Plans', 'Add-a-Line',
                   'Activations', 'Devices', 'Services', 'Accessories', 'Postpaid Plans', 'Prepaid Plans', 'eGAs',
                   'Store Locator Visits', 'A Actions', 'B Actions', 'C Actions', 'D Actions', 'E Actions', 'F Actions',
                   'Awareness Actions', 'Consideration Actions', 'Traffic Actions', 'Post-Click Activity',
                   'Post-Impression Activity', 'Video Views', 'Video Completions', 'Prepaid GAs', 'Postpaid GAs',
                   'Postpaid Orders', 'Prepaid Orders', 'Prepaid SIMs', 'Postpaid SIMs', 'Prepaid Mobile Internet',
                   'Postpaid Mobile Internet', 'Prepaid Phone', 'Postpaid Phone', 'Total GAs', 'DDR New Devices',
                   'DDR Add-a-Line']

        columns = dimensions + metrics + cfv_floodlight_columns + tag_columns

        data = data[columns]
        data.fillna(0, inplace=True)

        DataMethods().merge_past_data(data, columns, self.path)

        QA().placements(data)

    def metro(self):
        self.wb.save()

        sa = DataMethods().read_site_activity_report(adv='metro')
        cfv = DataMethods().read_cfv_report()

        data = sa.append(cfv)
        data = Floodlights().a_e_traffic(data, adv='metro')
        data = Categorization().date_columns(data)
        data = Categorization().categorize_report(data, adv='metro')

        data = self.additional_columns(data, adv='metro')

        sa_columns = list(sa.columns)
        tag_columns = sa_columns[sa_columns.index('Clicks') + 1:]

        dimensions = ['Week', 'Date', 'Month', 'Quarter', 'Campaign', 'Language', 'Site (DCM)', 'Site',
                      'TMO_Category', 'TMO_Category_Adjusted', 'Creative', 'Creative Type', 'Creative Groups 1',
                      'Creative ID', 'Ad', 'Creative Groups 2', 'Creative Field 2', 'Placement', 'Placement ID',
                      'Category', 'Creative Type Lookup', 'Skippable']

        cfv_floodlight_columns = ['Floodlight Attribution Type', 'Activity', 'Transaction Count']

        metrics = ['Media Cost', 'NTC Media Cost', 'Impressions', 'Clicks', 'Orders', 'Store Locator Visits',
                   'GM A Actions', 'GM B Actions', 'GM C Actions', 'GM D Actions', 'Hispanic A Actions',
                   'Hispanic B Actions', 'Hispanic C Actions', 'Hispanic D Actions', 'Total A Actions',
                   'Total B Actions', 'Total C Actions', 'Total D Actions', 'Awareness Actions',
                   'Traffic Actions', 'Consideration Actions', 'Post-Click Activity', 'Post-Impression Activity',
                   'Video Views', 'Video Completions']

        columns = dimensions + metrics + cfv_floodlight_columns + tag_columns

        data = data[columns]
        data = data.fillna(0)

        DataMethods().merge_past_data(data, columns, self.path)

        QA().placements(data)

    def dr_brand_remessaging(self):
        wb = Workbook.caller()
        wb.save()

        sa = DataMethods().read_site_activity_report(adv='dr')
        cfv2 = DataMethods().read_cfv_report()

        date = sa['Date'].max().strftime('%m.%d.%Y')

        feed_path = Range('Action_Reference', 'AE1').value

        cfv = Floodlights().custom_variables(cfv2)
        cfv = Floodlights().ddr_custom_variables(cfv)

        data = sa.append(cfv)
        data = self.strip_clickthroughs(data)

        data = Floodlights().a_e_traffic(data)

        data = Categorization().sites(data)
        data = Categorization().date_columns(data)
        data = Categorization().dr_placement_message_type(data)
        data = Categorization().dr_tactic(data)
        data = Categorization().placements(data)
        data = Categorization().dr_creative_categories(data)
        data = self.additional_columns(data, adv='dr')

        cfv_floodlight_columns = ['Activity', 'OrderNumber (string)', 'Plan (string)', 'Device (string)',
                                  'Service (string)', 'Accessory (string)', 'Floodlight Attribution Type', 'Orders',
                                  'Total GAs', 'Prepaid GAs', 'Postpaid GAs', 'Prepaid Orders', 'Postpaid Orders',
                                  'Prepaid SIMs', 'Postpaid SIMs', 'Prepaid Mobile Internet',
                                  'Postpaid Mobile Internet', 'Prepaid Phone', 'Postpaid Phone', 'DDR Add-a-Line',
                                  'DDR New Devices']

        sa_columns = list(sa.columns)
        tag_columns = sa_columns[sa_columns.index('DBM Cost (USD)') + 1:]

        dimensions1 = ['Campaign', 'Month', 'Week', 'Site', 'Tactic', 'Category', 'Placement Messaging Type',
                       'Message Bucket', 'Message Category', 'Message Offer']

        dimensions2 = ['Campaign2', 'Date', 'Site (DCM)', 'Creative', 'Click-through URL', 'Creative Pixel Size',
                       'Creative Type', 'Creative Field 1', 'Ad', 'Creative Groups 2', 'Placement', 'Placement ID',
                       'Placement Cost Structure']

        metrics1 = ['A Actions', 'B Actions', 'C Actions', 'D Actions', 'Store Locator Visits', 'Awareness Actions',
                    'Consideration Actions', 'Traffic Actions', 'Post-Impression Activity', 'Post-Click Activity',
                    'NTC Media Cost', 'NET Media Cost']

        metrics2 = ['Impressions', 'Clicks', 'Media Cost', 'DBM Cost (USD)']

        columns = dimensions1 + metrics1 + dimensions2 + metrics2 + tag_columns + cfv_floodlight_columns

        data = data[columns]

        data.fillna(0, inplace=True)

        DataMethods().merge_past_data(data, columns, wb.fullname)

        wb2 = Workbook()
        wb2.set_current()

        UpdateAdvertisers().top_15_devices(cfv2, feed_path)

        wb2.save(r'S:\SEA-Media\Analytics\T-Mobile\DR\Top 15 Devices Report\Top Devices Report ' + date + '.xlsx')
        wb2.close()

        wb.set_current()

    def search(self):
        sheets = ['DoubleClick Query', 'Google Location Intent Query', 'Bing Location Intent', 'Yelp', 'Whistleout',
                  'AdMarketplace', 'Marchex RAW', 'DFA Undefined and View Through']

        doubleclick = {
            'From': 'Day',
            'Account-Name': 'Account',
            'Impr': 'Impressions',
            'Cost': 'Net Spend',
            'Store Locator': 'Location Intent'
        }

        third_party = '|'.join(list(['Yelp', 'Whistleout', 'AdMarketplace']))

        google_location = {
            'From': 'Day',
            'Account-Name': 'Account',
            'Impr': 'Impressions',
            'Cost': 'Net Spend',
            'Clicks': 'Location Intent'
        }

        bing_location = {
            'Gregorian date': 'Day',
            'Account name': 'Account',
            'Impr': 'Impressions',
            'Cost': 'Net Spend',
            'Clicks': 'Location Intent',
            'Device type': 'Device segment'
        }

        yelp_whistleout = {
            'Date': 'Day',
            'Cost': 'Net Spend'
        }

        admarketplace = {
            'From': 'Day',
            'Spend': 'Net Spend'
        }

        marchex_raw = {
            'Date': 'Day',
            'Day of Time Filter': 'Day',
            'Account Name': 'Account',
            'Estimated Telesales Orders': 'Telesales Orders',
            'Estimated Gross Adds': "Telesales eGA's"
        }

        dfa_undefined = {
            'Paid Search Engine Account': 'Account',
            'Device Count (number)': "Online eGA's",
            'Order Count (number)': "Online Orders"
        }

        dc = pd.read_excel(self.path, sheets[0], index_col=None)
        del dc['Account']
        for col, renamed in doubleclick.iteritems():
            if col in dc.columns:
                dc.rename(columns={col: renamed}, inplace=True)
        dc['Engine'] = np.where(dc['Account'].str.contains(third_party) == True, '3rd Party', dc['Engine'])

        google = pd.read_excel(self.path, sheets[1], index_col=None)
        del google['Account']
        google['Engine'] = 'Google Adwords'
        for col, renamed in google_location.iteritems():
            if col in google.columns:
                google.rename(columns={col: renamed}, inplace=True)

        del google['Net Spend']
        del google['Impressions']

        bing = pd.read_excel(self.path, sheets[2], index_col=None)
        for col, renamed in bing_location.iteritems():
            if col in bing.columns:
                bing.rename(columns={col: renamed}, inplace=True)

        bing['Engine'] = 'Bing Ads'

        bing['Device segment'] = np.where(bing['Device segment'] == 'Computer', 'Desktop',
                                          np.where(bing['Device segment'] == 'Smartphone', 'Mobile',
                                                   bing['Device segment']))

        bing['Account'] = np.where(bing['Account'].str.contains('T-Mobile_') == True,
                                   bing['Account'].str.replace('T-Mobile_', ''), bing['Account'])

        bing['Brand / Generic / Comparison Shopping'] = \
            np.where(bing['Campaign name'].str.contains('_B_') == True, 'Brand',
                     np.where(bing['Campaign name'].str.contains('_G_') == True, 'Generic', None))

        del bing['Impressions']
        del bing['Net Spend']

        yelp = pd.read_excel(self.path, sheets[3], index_col=None)
        for col, renamed in yelp_whistleout.iteritems():
            if col in yelp.columns:
                yelp.rename(columns={col: renamed}, inplace=True)

        yelp['Account'] = 'Yelp'
        yelp['Engine'] = '3rd Party'
        yelp['Brand / Generic / Comparison Shopping'] = 'Brand'
        yelp['Device segment'] = 'Desktop'

        whistleout = pd.read_excel(self.path, sheets[4], index_col=None)
        for col, renamed in yelp_whistleout.iteritems():
            if col in whistleout.columns:
                whistleout.rename(columns={col: renamed}, inplace=True)

        whistleout['Account'] = 'Whistleout'
        whistleout['Engine'] = '3rd Party'
        whistleout['Brand / Generic / Comparison Shopping'] = 'Brand'
        whistleout['Device type'] = 'Desktop'

        admarket = pd.read_excel(self.path, sheets[5], index_col=None)
        for col, renamed in admarketplace.iteritems():
            if col in admarket.columns:
                admarket.rename(columns={col: renamed}, inplace=True)

        admarket['Account'] = 'AdMarketplace'
        admarket['Engine'] = '3rd Party'
        admarket['Brand / Generic / Comparison Shopping'] = 'Brand'
        admarket['Device type'] = 'Desktop'

        marchex = pd.read_excel(self.path, sheets[6], index_col=None)
        for col, renamed in marchex_raw.iteritems():
            if col in marchex.columns:
                marchex.rename(columns={col: renamed}, inplace=True)

        brand = '|'.join(list(['T-mobile - Hispanic Search', 'T-mobile - Search']))
        marchex['Account'] = np.where(marchex['Account'].str.contains(brand) == True, 'Brand', marchex['Account'])
        marchex['Account'] = marchex['Account'].apply(lambda x: x.replace('T-Mobile - ', ''))
        marchex['Engine'] = 'Marchex'
        marchex['Device type'] = 'Desktop'
        marchex[['Engine', 'Device segment']] = marchex['Group Name'].str.split(' ', expand=True)

        marchex['Engine'] = np.where(marchex['Engine'].str.contains('Google') == True, 'Google Adwords',
                                     np.where(marchex['Engine'].str.contains('Bing') == True, 'Bing Ads',
                                              marchex['Engine']))

        marchex['Engine'] = np.where(marchex['Engine'] == '3rd', '3rd Party', marchex['Engine'])
        marchex['Device segment'] = np.where(marchex['Device segment'] == 'Party', 'Desktop', marchex['Device segment'])

        dfa = pd.read_excel(self.path, sheets[7], index_col=None)
        for col, renamed in dfa_undefined.iteritems():
            if col in dfa.columns:
                dfa.rename(columns={col: renamed}, inplace=True)

        vals_to_replace = '|'.join(list(['T-Mobile Bing ', 'T-Mobile Google ', 'DART Search : ', 'DART Search: ']))
        dfa['Account'] = dfa['Placement'].astype(str).apply(lambda x: x.replace(vals_to_replace, ''))
        # dfa['Account'] = dfa['Account'].astype(str).apply(lambda x: x.replace('T-Mobile Google ', ''))
        # dfa['Account'].fillna('Whistleout', inplace=True)
        # dfa['Account'] = np.where(pd.isnull(dfa['Account']) == False, 'Whistleout', dfa['Account'])

        dfa['Device segment'] = 'Desktop'

        # dfa['Engine'] = dfa['Site (DCM)'].apply(lambda x: x.replace({engines}, regex=True))
        dfa['Engine'] = np.where(dfa['Site (DCM)'] == 'DART Search: Whistleout', '3rd Party',
                                 np.where(dfa['Site (DCM)'] == 'DART Search : Google', 'Google Adwords',
                                          np.where(dfa['Site (DCM)'] == 'DART Search : MSN', 'Bing Ads', None)))

        # dfa['Engine'] = dfa['Site (DCM)'].apply(lambda x: x.replace('DART Search : ', ''))

        # dfa['Engine'] = dfa['Site (DCM)'].str.replace('DART Search: ', '')
        # dfa['Engine'] = dfa['Site (DCM)'].str.replace('DART Search : ', '')

        # dfa['Engine'] = np.where(dfa['Engine'] == 'Whistleout', '3rd Party', np.where(dfa['Engine'] == 'Google', 'Google Adwords',
        #                                                                              np.where(dfa['Engine'] == 'MSN', 'Bing Ads', dfa['Engine'])))

        dfa["Online eGA's"] = np.where(dfa['ORD Value'] == 'undefined', 1, dfa["Online eGA's"])

        dfa['Brand / Generic / Comparison Shopping'] = np.where(dfa['Paid Search Campaign'].str.contains('_B_') == True,
                                                                'Brand',
                                                                np.where(dfa['Paid Search Campaign'].str.contains(
                                                                    '_G_') == True, 'Generic',
                                                                         np.where(
                                                                             dfa['Paid Search Campaign'].str.contains(
                                                                                 'Shop') == True, 'Comparison Shopping',
                                                                             'Brand')))

        dashboard_data = dc.append([google, bing, yelp, whistleout, admarket, marchex, dfa])

        dashboard_data['Date2'] = pd.to_datetime(dashboard_data['Day'])
        dashboard_data['Month'] = dashboard_data['Date2'].fillna(0).apply(lambda x: arrow.get(x).format('MMMM'))
        dashboard_data['Week'] = dashboard_data['Date2'].fillna(0).apply(lambda x: Categorization().mondays(x))
        dashboard_data.drop('Date2', axis=1, inplace=True)

        dashboard_data = dashboard_data[['Month', 'Week', 'Day', 'Account', 'Engine',
                                         'Brand / Generic / Comparison Shopping', 'Campaign', 'Device segment',
                                         'Net Spend', 'Impressions', 'Clicks', 'Location Intent', 'Online Orders',
                                         'Telesales Orders', "Online eGA's", "Telesales eGA's"]]

        DataMethods().chunk_df(dashboard_data, 'output', 'A1')

    def wfm(self):
        def output_columns(data):
            data_columns = list(data.columns)
            dimensions = ['Campaign', 'Date', 'Week', 'Month', 'Quarter', 'Site (DCM)', 'Site', 'Creative', 'Ad',
                          'Creative Groups 1', 'Creative Groups 2', 'Creative Field 1', 'Placement', 'Placement ID',
                          'Placement Cost Structure']

            metrics = ['Media Cost', 'NTC Media Cost', 'Impressions', 'Clicks', 'Traffic Actions Hub',
                       'Traffic Actions Walmart', 'Video Completions', 'Video Views']

            floodlight_columns = data_columns[data_columns.index('Clicks') + 1:]

            columns = dimensions + metrics + floodlight_columns

            return list(columns)

        def wfm_columns(data):
            column_names = list(data.columns)
            hub_key = '|'.join(list(['WFM', 'Hub']))
            walmart_key = '|'.join(list(['Walmart Family Mobile', 'Walmart']))

            hub_actions, walmart_actions = [], []

            for i in column_names:
                hub = re.search(hub_key, i)
                walmart = re.search(walmart_key, i)
                if hub:
                    hub_actions.append(i)
                if walmart:
                    walmart_actions.append(i)

            hub_actions = list(set(hub_actions).intersection(column_names))
            walmart_actions = list(set(walmart_actions).intersection(column_names))

            data['Traffic Actions Hub'] = data[hub_actions].sum(axis=1)
            data['Traffic Actions Walmart'] = data[walmart_actions].sum(axis=1)

            data['Video Completions'] = 0
            data['Video Views'] = 0

            data['NTC Media Cost'] = data['Media Cost'] / .96759

            return data

        self.wb.save()

        data = pd.read_excel(self.path, self.temp_site_activity, index_col=None)
        data = Categorization().sites(data)
        data = Categorization().placements(data, adv='wfm')
        data = Categorization().date_columns(data)
        data = wfm_columns(data)

        ordered_columns = output_columns(data)

        data = data[ordered_columns]

        if Range('data', 'A1').value is None:
            DataMethods().chunk_df(data, 'data', 'A1')

        # If data is already present in the tab, the two data sets are merged together and then copied into the data tab.

        else:
            past_data = pd.read_excel(self.path, 'data', index_col=None)
            appended_data = past_data.append(data)
            appended_data = appended_data[ordered_columns]
            appended_data.fillna(0, inplace=True)
            Sheet('data').clear()
            DataMethods().chunk_df(appended_data, 'data', 'A1')

    @staticmethod
    def top_15_devices(cfv, feed_path):
        Sheet.add('DDR')
        Sheet.add('Summary')

        device_text_file = pd.read_table(feed_path)

        excluded = Range('Lookup', 'O2').value

        cfv['Device IDs'] = cfv['Device (string)'].str.split(',')

        cfv['Plan Names'] = cfv['Plan (string)'].str.split(',')

        ddr_devices = pd.Series(list(np.where((cfv['Campaign'].str.contains('DDR') == True) |
                                              (cfv['Campaign'].str.contains('Brand Remessaging') == True),
                                              cfv['Device IDs'], np.NaN)))

        ddr_plans = pd.Series(list(np.where((cfv['Campaign'].str.contains('DDR') == True) |
                                            (cfv['Campaign'].str.contains('Brand Remessaging') == True),
                                            cfv['Plan Names'], np.NaN)))

        ddr_devices.dropna(inplace=True)
        ddr_plans.dropna(inplace=True)

        ddr_devices = list(itertools.chain(*ddr_devices))
        ddr_plans = list(itertools.chain(*ddr_plans))

        while '' in ddr_devices: ddr_devices.remove('')
        while '' in ddr_plans: ddr_plans.remove('')
        while excluded in ddr_devices: ddr_devices.remove(excluded)

        device_counts = pd.DataFrame(pd.value_counts(pd.Series(ddr_devices).values, sort=True)[0:15])
        device_counts['Device Name'] = 1

        plan_counts = pd.DataFrame(pd.value_counts(pd.Series(ddr_plans).values, sort=True)[0:15])

        Range('DDR', 'A1', index=False).value = device_text_file

        Range('Summary', 'B1').value = device_counts

        Range('Summary', 'I1').value = plan_counts

        Sheet('Summary').activate()

        # Rank

        i = 0
        for cell in Range('Summary', 'A2:' + 'A' + str(len(device_counts) + 1)):
            i += 1
            cell.value = i

        j = 0
        for cell in Range('Summary', 'H2:' + 'H' + str(len(plan_counts) + 1)):
            j += 1
            cell.value = j

        # Device Name

        for cell in Range('Summary', 'D2').vertical:
            ids = cell.offset(0, -2).get_address(False, False, False)
            cell.formula = '=IFERROR(INDEX(DDR!A:A,MATCH(Summary!' + ids + ',DDR!G:G,0)),"na")'

        Range('Summary', 'A1:C1').value = 'Rank', 'Device SKU', 'Count'
        Range('Summary', 'H1').value = 'Rank'
        Range('Summary', 'I1').value = 'Plan Name'

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
        data['Click-through URL'] = data['Click-through URL'].str.replace(
            '%3F%3DADV_DS_%epid!_%eaid!_%ecid!_%eadv!',
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