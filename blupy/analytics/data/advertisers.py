import itertools
import re

import numpy as np
import pandas as pd
import arrow
from xlwings import Workbook, Range, Sheet

from analytics.reporting.qa import QA
from categorization import Categorization
from io import DataMethods
from floodlights import Floodlights


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
        data = DataMethods().strip_clickthroughs(data)

        data = Floodlights().a_e_traffic(data, adv='tmo')

        data = Categorization().categorize_report(data, adv='tmo')
        data = Floodlights().f_tags(data)
        data = DataMethods().additional_columns(data, adv='tmo')

        sa_columns = list(sa.columns)
        tag_columns = sa_columns[sa_columns.index('DBM Cost (USD)') + 1:]

        columns = DataMethods().order_columns(adv='tmo') + tag_columns

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

        data = DataMethods().additional_columns(data, adv='metro')

        sa_columns = list(sa.columns)
        tag_columns = sa_columns[sa_columns.index('Clicks') + 1:]

        columns = DataMethods().order_columns(adv='metro') + tag_columns

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
        data = DataMethods().strip_clickthroughs(data)

        data = Floodlights().a_e_traffic(data)

        data = Categorization().sites(data)
        data = Categorization().date_columns(data)
        data = Categorization().dr_placement_message_type(data)
        data = Categorization().dr_tactic(data)
        data = Categorization().placements(data)
        data = Categorization().dr_creative_categories(data)
        data = DataMethods().additional_columns(data, adv='dr')

        cfv_floodlight_columns = ['Activity', 'OrderNumber (string)', 'Plan (string)', 'Device (string)',
                                  'Service (string)', 'Accessory (string)', 'Floodlight Attribution Type', 'Orders',
                                  'Total GAs', 'Prepaid GAs', 'Postpaid GAs', 'Prepaid Orders', 'Postpaid Orders',
                                  'Prepaid SIMs', 'Postpaid SIMs', 'Prepaid Mobile Internet',
                                  'Postpaid Mobile Internet', 'Prepaid Phone', 'Postpaid Phone', 'DDR Add-a-Line',
                                  'DDR New Devices']

        sa_columns = list(sa.columns)
        tag_columns = sa_columns[sa_columns.index('DBM Cost (USD)') + 1:]

        columns = DataMethods().order_columns(adv='dr') + tag_columns + cfv_floodlight_columns

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