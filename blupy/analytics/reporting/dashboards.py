import string

import numpy as np
import pandas as pd
from xlwings import Range, Workbook, Sheet

from analytics.data.file_io import DataMethods


class DrDashboard(DataMethods):

    def __init__(self):
        super(DrDashboard, self).__init__()
        self.save_path = self.path[:self.path.rindex('\\')]

    def generate_dashboard_data(self):
        wb = Workbook.caller()

        path = Range('Sheet3', 'AB1').value

        ddr_data = pd.read_excel(path, 'data', index_cols=None)
        ddr_data.fillna(0, inplace=True)

        ddr_display = DrDashboard().display_data(ddr_data)
        ddr_search_data = DrDashboard().merge_search_data()

        tableau_search = DrDashboard().search_data(ddr_search_data)

        tableau = ddr_display.append(tableau_search)

        tableau['Quarter'] = None
        tableau['Week of Quarter'] = None

        if Range('merged', 'A1').value is None:
            DataMethods().chunk_df(tableau, 'merged', 'A1')

        # If data is already present in the tab, the two data sets are merged together and then copied into the data tab.
        else:
            past_data = pd.read_excel(wb.fullname, 'merged', index_col=None)
            past_data = past_data[past_data['Campaign'] != 'Search']
            past_data = past_data[(past_data['Source'] == 'DR-Pivot') & (past_data['Week'] <= '12/31/2015')]
            appended_data = past_data.append(tableau)
            Sheet('merged').clear()
            DataMethods().chunk_df(appended_data, 'merged', 'A1')

        DrDashboard().search_data_client(ddr_search_data)
        DrDashboard().display_data_client(ddr_data)

    @staticmethod
    def display_data(ddr):
        ddr_columns = ['Campaign', 'Week', 'Site', 'Message Tactic', 'Placement Messaging Type', 'A Actions',
                       'B Actions',
                       'C Actions', 'D Actions', 'Store Locator Visits', 'Awareness Actions', 'Consideration Actions',
                       'Traffic Actions', 'View-through Conversions', 'Click-through Conversions', 'NTC Media Cost',
                       'NET Media Cost', 'Impressions', 'Clicks', 'Orders', 'Prepaid Orders', 'Postpaid Orders',
                       'Total GAs', 'Prepaid GAs', 'Postpaid GAs', 'Prepaid SIMs', 'Postpaid SIMs',
                       'Prepaid Mobile Internet', 'Postpaid Mobile Internet', 'Prepaid Phone', 'Postpaid Phone',
                       'DDR Add-a-Line', 'DDR New Devices']

        ddr.rename(columns={'Tactic': 'Message Tactic',
                            'Post-Impression Activity': 'View-through Conversions',
                            'Post-Click Activity': 'Click-through Conversions'}, inplace=True)

        ddr_data = ddr[ddr_columns]

        telesales = pd.DataFrame(Range('Telesales', 'A1').table.value,
                                 columns=Range('Telesales', 'A1').horizontal.value)
        telesales.drop(0, inplace=True)
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

    @staticmethod
    def merge_search_data():
        sheet = Range('Sheet3', 'AC1').value

        search_tabs = ['Search Raw Data', 'Whistleout', 'CSE', 'Ad Marketplace']

        search_data = pd.DataFrame()

        for i in search_tabs:
            df = pd.read_excel(sheet, i, index_cols=None)
            df['Source'] = i
            search_data = search_data.append(df)

        return search_data

    @staticmethod
    def search_data(search_data):
        search_data[['Total GAs', 'New Total eGAs']] = search_data[['Total GAs', 'New Total eGAs']].fillna(0)

        search_data['Total GAs'] = search_data['Total GAs'] + search_data['New Total eGAs']
        search_data['NTC Media Cost'] = search_data['NET Media Cost'] / .96759

        search_gas = pd.DataFrame(Range('Search_GAs', 'A1').table.value,
                                  columns=Range('Search_GAs', 'A1').horizontal.value)
        search_gas.drop(0, inplace=True)

        search_gas = search_gas[['Week', 'Traffic Actions', 'Prepaid GAs', 'Postpaid GAs', 'Prepaid SIMs',
                                 'Postpaid SIMs']]

        search_gas['Source'] = 'Search Raw Data'

        search = search_data.append(search_gas)

        search['Tactic'] = 'Search'
        search['Channel'] = 'DR'
        search['Campaign'] = 'Search'

        return search

    def search_data_client(self, search_data):
        column_order = ['Search Engine', 'Brand DDR Bucket', 'Week', 'NET Media Cost', 'Impressions', 'Clicks',
                        'Orders',
                        'Plans', 'Prepaid Orders', 'Consideration Actions', 'Add-A-Line', 'Total GAs', 'New Total eGAs',
                        'Telesales GAs']

        client_data = search_data[column_order]
        client_data = client_data[client_data['Week'] >= '1/1/2016']

        client_data['Brand DDR Bucket'] = np.where(pd.isnull(client_data['Brand DDR Bucket']) == True,
                                                   client_data['Search Engine'],
                                                   client_data['Brand DDR Bucket'])

        client_data.to_csv(self.save_path + '\\' + 'DR_Search_Raw_Data.txt', sep='\t', encoding='utf-8', index=False)

    def display_data_client(self, dr_data):
        dr_data.to_csv(self.save_path + '\\' + 'DDR_Raw_Data.txt', sep='\t', encoding='utf-8', index=False)


class CrossChannel(DataMethods):

    def __init__(self):
        super(CrossChannel, self).__init__()

    @staticmethod
    def dma_lookup():
        dma = pd.read_excel(Range('Ref', 'B1').value, 0, index_col= 'DMA')

        return dma

    @staticmethod
    def competitor_data(dma, path):
        dropped_columns = ['CATEGORY', 'MICROCATEGORY', 'PARENT', 'BRAND', 'TITLE', 'ADSCOPE ID']
        renamed_columns = {'Time Period':'Week', 'Length/size':'Creative Size', 'Market':'DMA', 'Dols (000)':'Spend'}

        data = pd.read_excel(path, index_col=None)

        for col in data.columns:
            if col in dropped_columns:
                data.drop(col, axis=1, inplace=True)
            data.rename(columns={col : string.capwords(col)}, inplace=True)

        for col, renamed in renamed_columns.iteritems():
            if col in data.columns:
                data.rename(columns={col : renamed}, inplace=True)


        data['Week'] = pd.to_datetime(data['Week'].apply(lambda x: x.split(' ')[1]))
        data['Spend'] = data['Spend'] * 1000

        data['Subcategory'] = np.where(data['Subcategory'].str.contains('Consumer Wireless') == True, 'Consumer Wireless',
                                       np.where(data['Subcategory'].str.contains('Business Wireless') == True, 'Business Wireless',
                                                np.where(data['Subcategory'].str.contains('Pre-Paid') == True, 'Pre-Paid', None)))

        data['DMA'] = data['DMA'].apply(lambda x: string.capwords(x))
        data['DMA'] = data['DMA'].apply(lambda x: x.replace('* National', 'National'))
        data.set_index('DMA', inplace=True)

        data = pd.merge(data, dma, how='left', right_index=True, left_index=True).reset_index()
        data_pivoted = pd.pivot_table(data, index=['Advertiser', 'Week', 'Subcategory', 'Media', 'DMA'],
                                      aggfunc=np.sum).reset_index()
        data_pivoted.rename(columns={'Subcategory':'Category'}, inplace=True)

        return data_pivoted

    @staticmethod
    def offline_data(dma, path, branded_search_impressions):
        tabs = {'nat_tv':'National TV', 'ooh':'Out of Home', 'newspaper':'Newspaper', 'radio':'Radio', 'spot_tv':'Spot TV'}
        offline_df = pd.DataFrame()

        for tab, medium in tabs.iteritems():
            df = pd.read_excel(path, tab, index_col=None)
            df['Medium'] = medium
            if tab == 'radio':
                df['Medium'] = df['Network/Spot']
                df.drop('Network/Spot', axis=1, inplace=True)
            if tab == 'nat_tv':
                df.set_index('Week', inplace=True)
                df = pd.merge(df, branded_search_impressions, how='left', right_index=True, left_index=True)
                df.reset_index(inplace=True)
            offline_df = offline_df.append(df)

        weeks = list(offline_df[offline_df['Medium'] == 'National TV']['Week'].unique())

        for i in weeks:
            offline_df['Branded Search Impressions'] = np.where(
                (offline_df['Week'] == i) & (offline_df['Medium'] == 'National TV'),
                offline_df['Branded Search Impressions'] / len(
                    offline_df[(offline_df['Week'] == i) & (offline_df['Medium'] == 'National TV')]),
                offline_df['Branded Search Impressions'])

        offline_df.set_index('DMA', inplace=True)
        offline_df = pd.merge(offline_df, dma, how='left', right_index=True, left_index=True).reset_index()

        return offline_df

    @staticmethod
    def online_data(path):
        consolidated = pd.read_excel(path, 'data', index_col=None,
                                     parse_cols='A:BJ')

        needed_columns = ['Week', 'Media Plan', 'Language', 'Site', 'Creative Groups 2', 'Sub-Tactic', 'Tactic',
                          'NTC Media Cost', 'Impressions', 'Clicks', 'Orders', 'Traffic Actions', 'Video Views',
                          'Video Completions']

        for col in consolidated.columns:
            if col not in needed_columns:
                consolidated.drop(col, axis=1, inplace=True)

        consolidated.rename(columns={'Creative Groups 2':'Message', 'NTC Media Cost':'Spend',
                                     'Site':'Publisher', 'Media Plan':'Campaign'}, inplace=True)

        consolidated_pivot = pd.pivot_table(consolidated, index=['Week', 'Campaign', 'Language', 'Publisher', 'Message',
                                                                 'Tactic', 'Sub-Tactic'], aggfunc=np.sum).reset_index()

        consolidated_pivot['Medium'] = 'Online'

        return consolidated_pivot

    @staticmethod
    def social_data(path):
        soc = pd.read_excel(path, 0, index_col=None)

        unneeded_columns = ['Ad ID', 'Camp ID', 'Tweet ID']

        for col in soc.columns:
            if col in unneeded_columns:
                soc.drop(col, axis=1, inplace=True)
            if ' - All' in col:
                soc.rename(columns={col : col.replace(' - All', '')}, inplace=True)

        renamed = {'NTC Spend':'Spend', 'Ad Video P100 Watched':'Video Completions',
                   'Site':'Publisher', 'Week Starting':'Week'}

        for col, renamed_col in renamed.iteritems():
            if col in soc.columns:
                soc.rename(columns={col : renamed_col}, inplace=True)

        #soc['Impressions'] = soc[['Qualified Impressions', 'Impressions']].sum(axis=1)
        soc['Language'] = np.where(soc['Language'].str.contains('English') == True, 'EL', 'SL')

        soc_pivot = pd.pivot_table(soc, index=['Week', 'Campaign', 'Campaign Objective', 'Creative Type', 'Publisher',
                                               'Language', 'Interest']).reset_index()

        soc_pivot['Medium'] = 'Social'

        return soc_pivot

    @staticmethod
    def search_data(path):
        data = pd.read_excel(path, 0, index_col=None)
        cols_dropped = ['Row Type', 'To', 'Engine Status', 'Status', 'Sync errors', 'CTR', 'Avg CPC', 'Avg pos',
                        'Daily budget']
        search_rename = {'From':'Week', 'Impr':'Impressions', 'Cost':'Spend'}

        for col in data.columns:
            if col in cols_dropped:
                data.drop(col, axis=1, inplace=True)

        for col, renamed_col in search_rename.iteritems():
            if col in data.columns:
                data.rename(columns={col : renamed_col}, inplace=True)

        data['Branded Search Impressions'] = np.where(data['Campaign'].str.contains('_B_') == True, data['Impressions'], 0)

        data_pivoted = pd.pivot_table(data, index=['Week', 'Engine', 'Account'], values=['Clicks', 'Impressions', 'Spend',
                                                                                         'Branded Search Impressions'],
                                      aggfunc=np.sum).reset_index()

        data_pivoted['Medium'] = 'Search'

        return data_pivoted

    @staticmethod
    def tmo_inputs(path):
        inputs = ['Traffic', 'Credit Apps', 'Gross Adds', 'Retail Traffic']

        input_data = pd.DataFrame()

        for i in inputs:
            df = pd.read_excel(path, i, index_col=None)

            if i == 'Credit Apps':
                df.rename(columns={i :'Volume'}, inplace=True)
                df['Metric'] = i

            if i == 'Gross Adds':
                df.rename(columns={i :'Volume'}, inplace=True)
                df['Metric'] = i

            if i == 'Retail Traffic':
                df.rename(columns={i :'Volume'}, inplace=True)
                df['Metric'] = i

            if i == 'Traffic':
                df['Total Traffic'] = df['Customer Traffic'] + df['Prospect Traffic']
                df = pd.melt(df, id_vars=['Week'], var_name='Metric', value_name='Volume')

            input_data = input_data.append(df)

        return input_data

    @staticmethod
    def merge_data(online, offline, search, social, tmo):
        online_combined = online.append([social, search])

        online_combined['Sub-Tactic'] = np.where(online_combined['Sub-Tactic'].isnull() == True,
                                                 online_combined['Medium'],
                                                 online_combined['Sub-Tactic'])
        online_combined['Tactic'] = np.where(online_combined['Tactic'].isnull() == True, online_combined['Medium'],
                                             online_combined['Tactic'])

        tmo_inputs = pd.pivot_table(tmo, index=['Week'], columns=['Metric'], values=['Volume'], aggfunc=np.sum)
        tmo_inputs.columns = tmo_inputs.columns.get_level_values(1)

        search_impressions = pd.pivot_table(search, index=['Week'], values=['Branded Search Impressions'],
                                            aggfunc=np.sum).reset_index()

        online_offline = online_combined.append(offline)
        online_offline.set_index('Week', inplace=True)

        online_offline_tmo = pd.merge(online_offline, tmo_inputs, how='left',
                                      right_index=True, left_index=True).reset_index()

        weeks = list(online_offline_tmo['Week'].unique())
        metric_columns = ['Customer Traffic', 'Direct Load Traffic', 'Gross Adds', 'Mobile Visits',
                          'Non-Mobile Visits', 'Prospect Traffic', 'Retail Traffic', 'Total Orders', 'Total Traffic']

        for i in weeks:
            for j in metric_columns:
                online_offline_tmo[j] = np.where(online_offline_tmo['Week'] == i,
                                                 online_offline_tmo[j] / len(
                                                     online_offline_tmo[online_offline_tmo['Week'] == i]),
                                                 online_offline_tmo[j])

        aggregated_tmo = pd.pivot_table(tmo, index=['Metric', 'Week'], aggfunc=np.sum).reset_index()
        search_impressions['Metric'] = 'Branded Search'
        search_impressions.rename(columns={'Branded Search Impressions': 'Volume'}, inplace=True)

        aggregated_tmo = aggregated_tmo.append(search_impressions)
        online_offline.reset_index(inplace=True)

        aggregated_tmo.set_index('Week', inplace=True)
        spend = pd.pivot_table(online_offline, index=['Week'], values=['Spend'], aggfunc=np.sum)
        aggregated_tmo = pd.merge(aggregated_tmo, spend, how='left', right_index=True, left_index=True).reset_index()

        weeks = list(aggregated_tmo['Week'].unique())

        for i in weeks:
            aggregated_tmo['Spend'] = np.where(aggregated_tmo['Week'] == i,
                                               aggregated_tmo['Spend'] / len(
                                                   aggregated_tmo[aggregated_tmo['Week'] == i]),
                                               aggregated_tmo['Spend'])

        return [online_offline_tmo, aggregated_tmo]

    @staticmethod
    def generate_dashboard():
        search_path, tmo_inputs_path, adscope_path, online_path, social_path, offline_path = Range('Ref', 'B2:B7').value

        dma = CrossChannel().dma_lookup()
        competitor = CrossChannel().competitor_data(dma, adscope_path)

        search_dat = CrossChannel().search_data(search_path)

        search_impressions = pd.pivot_table(search_dat, index=['Week'],
                                                    values=['Branded Search Impressions'], aggfunc=np.sum)

        data = CrossChannel().merge_data(CrossChannel().online_data(online_path), CrossChannel().offline_data(dma, offline_path,
                                                                                         search_impressions),
                                            search_dat, CrossChannel().social_data(social_path),
                                         CrossChannel().tmo_inputs(tmo_inputs_path))

        data_updates = {'Competitive' : competitor, 'merged_channels' : data[0], 'tmo_volume' : data[1]}

        for i, j in data_updates.iteritems():
            Sheet(i).clear_contents()
            DataMethods().chunk_df(j, i, 'A1')