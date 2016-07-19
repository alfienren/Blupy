import pandas as pd
import numpy as np
from xlwings import Range
import string


def dma_lookup():
    dma = pd.read_excel(Range('Ref', 'B1').value, 0, index_col= 'DMA')

    return dma


def competitors(dma, path):
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


def offline(dma, path, branded_search_impressions):
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


def online(path):
    consolidated = pd.read_excel(path, 'data', index_col=None,
                                 parse_cols='A:AE,AI:AL,BE,BH:BI')

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


def social(path):
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

    soc['Impressions'] = soc[['Qualified Impressions', 'Impressions']].sum(axis=1)
    soc['Language'] = np.where(soc['Language'].str.contains('English') == True, 'EL', 'SL')

    soc_pivot = pd.pivot_table(soc, index=['Week', 'Campaign', 'Campaign Objective', 'Creative Type', 'Publisher',
                                           'Language', 'Interest']).reset_index()

    soc_pivot['Medium'] = 'Social'

    return soc_pivot


def search(path):
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

    data['Branded Search Impressions'] = np.where(data['Campaign'].apply(lambda x: x[x.find('(') + 1:x.find(')')]) == 'B',
                                                  data['Impressions'], 0)

    data_pivoted = pd.pivot_table(data, index=['Week', 'Engine', 'Account'], values=['Clicks', 'Impressions', 'Spend',
                                                                                     'Branded Search Impressions'],
                                  aggfunc=np.sum).reset_index()

    data_pivoted['Medium'] = 'Search'

    return data_pivoted


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


