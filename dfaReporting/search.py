from xlwings import Range, Sheet, Workbook
import pandas as pd
import numpy as np
import os
from reporting import datafunc, custom_variables, categorization
from reporting.constants import TabNames


def query_and_cfv_data(path):
    query_sheets = ['Weekly Dash', 'Sitelink DDR', 'Sitelink Remarketing', 'Sitelink Bing',
                    'Retention Location Intent Query', 'Location Intent Query', 'Marchex']

    buckets = pd.read_excel(path, 'Lookup', parse_cols='E:G')

    search_data = pd.DataFrame()

    for i in query_sheets:
        data = pd.read_excel(path, i, index_col=None)

        if i == 'Sitelink Bing':
            dr_brand = '|'.join(list(['DDR_B_High_Volume', 'DDR_B_Bring']))
            remarket = '|'.join(list(['REM_']))
            b_marketing = '|'.join(list(['DDR_B_']))
            data.rename(columns={'Campaign name':'Campaign', 'Ad extension property value':'Sitelink display text',
                                 'Impressions': 'Impr', 'Spend': 'Cost'}, inplace=True)

            data = data[(data['Sitelink display text'].str.contains('My T-Mobile') == True)
                        | (data['Sitelink display text'].str.contains('for Business') == True)]

            data['Bucket Class'] =  np.where(data['Campaign'].str.contains(dr_brand) == True, 'DR-Brand',
                                             np.where(data['Campaign'].str.contains(remarket) == True, 'Remarketing',
                                                      np.where(data['Campaign'].str.contains(b_marketing) == True,
                                                               'Brand Marketing', None)))

        if i == 'Marchex':
            data.set_index(['campaign_bucket', 'Metric'], inplace=True)
            data = data.stack().reset_index()
            data = pd.pivot_table(data, index=['campaign_bucket', 'level_2'], columns='Metric',
                                  aggfunc=np.sum).reset_index()

            data = pd.merge(data[['campaign_bucket', 'level_2']], data[0],
                            how='left', right_index=True, left_index=True)

            data.rename(columns={data.columns[0]: 'Bucket', data.columns[1]: 'From'}, inplace=True)

            data['Bucket Class'] = np.where(data['Bucket'].str.contains('DR') == True, 'DR-Brand',
                                            np.where(data['Bucket'] == 'Brand Marketing', 'Brand Marketing',
                                                     np.where(data['Bucket'] == 'Remarketing', 'Remarketing',
                                                              np.where(data['Bucket'] == 'Deals & Coupons',
                                                                       'Deals/Coupons',
                                                                       np.where(data['Bucket'] == 'Whistleout',
                                                                                'Whistleout', None)))))

        data['Source'] = i
        search_data = search_data.append(data)

    search_data.rename(columns={'From':'Date'}, inplace=True)

    cfv = datafunc.read_cfv_report(path)

    buckets_cfv = buckets.set_index(['Campaign', 'Creative'])
    cfv.set_index(['Campaign', 'Creative'], inplace=True)
    cfv = pd.merge(cfv, buckets_cfv, how='left', right_index=True, left_index=True).reset_index()

    cfv = custom_variables.custom_variable_columns(cfv)
    cfv = custom_variables.ddr_custom_variables(cfv)

    cfv['Source'] = 'CFV'

    search_data = search_data.append(cfv)
    search_data['Date'] = pd.to_datetime(search_data['Date'])
    search_data['Week'] = search_data['Date'].apply(lambda x: categorization.mondays(x))

    search_data.fillna(0, inplace=True)

    search_pivoted = pd.pivot_table(search_data, index=['Source', 'Account', 'Bucket Class', 'Business Unit', 'Campaign',
                                          'Device segment', 'Date', 'Sitelink display text', 'Accessory (string)',
                                          'OrderNumber (string)', 'Device (string)', 'Plan (string)',
                                          'Service (string)'],
                             aggfunc=np.sum, fill_value=0).reset_index()

    search_pivoted['Location Intent_temp'] = np.where(search_pivoted['Source'] == 'Weekly Dash',
                                                      search_pivoted['Store Locator'], 0)
    search_pivoted['Location Intent_temp2'] = np.where(search_pivoted['Source'] == 'Location Intent Query',
                                                       search_pivoted['Clicks'], 0)

    search_pivoted['Location Intent'] = search_pivoted['Location Intent_temp'] + search_pivoted['Location Intent_temp2']
    search_pivoted.drop(['Location Intent_temp', 'Location Intent_temp2'], axis=1, inplace=True)

    search_pivoted.rename(columns={'Cost': 'Spend', 'Impr': 'Impressions'}, inplace=True)

    search_pivoted = categorization.search_bucket_class(search_pivoted)
    search_pivoted = categorization.date_columns(search_pivoted)

    datafunc.chunk_df(search_pivoted, 'data', 'A1')


def save_raw_data_file(search_pivoted):
    save_path = Range('Lookup', 'K1').value

    through_date = search_pivoted['Date'].max().strftime('%m.%d.%Y')

    file_name = 'Search_Raw_Data_' + through_date + '.xlsx'

    search_pivoted.to_excel(os.path.join(save_path, file_name), index=False)


def generate_search_reporting():
    wb = Workbook.caller()
    wb.save()

    query_and_cfv_data(wb.fullname)


def search_cfv_report(path):
    buckets = pd.read_excel(path, 'Lookup', parse_cols='E:G')

    cfv = datafunc.read_cfv_report(path)

    cfv = custom_variables.custom_variable_columns(cfv)
    cfv = custom_variables.ddr_custom_variables(cfv)

    buckets_cfv = buckets.set_index(['Campaign', 'Creative'])
    cfv.set_index(['Campaign', 'Creative'], inplace=True)
    cfv = pd.merge(cfv, buckets_cfv, how='left', right_index=True, left_index=True).reset_index()

    cfv = categorization.search_cfv_categories(cfv)

    return cfv


def search_cfv_outputs(cfv):
    bucket_class_table = pd.pivot_table(cfv, index=['Bucket Class'],
                                        values=['Prepaid Orders', 'Postpaid Orders', 'Orders',
                                                'Prepaid GAs', 'Postpaid GAs', 'Total GAs'],
                                        aggfunc=np.sum).reset_index()

    engine_web_team_orders_table = pd.pivot_table(cfv, index=['Web Team'],
                                           values=['Postpaid Orders', 'Prepaid Orders', 'Orders'],
                                           aggfunc=np.sum).reset_index()

    total_prepaid_orders, total_postpaid_orders, total_orders = [engine_web_team_orders_table['Prepaid Orders'].sum(),
                                                                 engine_web_team_orders_table['Postpaid Orders'].sum(),
                                                                 engine_web_team_orders_table['Prepaid Orders'].sum() +
                                                                 engine_web_team_orders_table['Postpaid Orders'].sum()]

    prepaid_order_percent = total_prepaid_orders / total_orders
    postpaid_order_percent = total_postpaid_orders / total_orders

    engine_web_team_ga_table = pd.pivot_table(cfv, index=['Web Team'],
                                                  values=['Postpaid GAs', 'Prepaid GAs', 'Total GAs'],
                                                  aggfunc=np.sum).reset_index()

    orders_offset = len(engine_web_team_orders_table)
    gas_offset = len(engine_web_team_ga_table)

    total_prepaid_gas, total_postpaid_gas, total_gas = [engine_web_team_ga_table['Prepaid GAs'].sum(),
                                                                 engine_web_team_ga_table['Postpaid GAs'].sum(),
                                                                 engine_web_team_ga_table['Prepaid GAs'].sum() +
                                                                 engine_web_team_ga_table['Postpaid GAs'].sum()]

    prepaid_ga_percent = total_prepaid_gas / total_gas
    postpaid_ga_percent = total_postpaid_gas / total_gas

    Sheet(TabNames.search_output).clear()

    Range(TabNames.search_output, 'A1', index=False).value = bucket_class_table
    Range(TabNames.search_output, 'K1', index=False).value = \
        engine_web_team_orders_table[['Web Team', 'Postpaid Orders', 'Prepaid Orders', 'Orders']]
    Range(TabNames.search_output, 'R1', index=False).value = engine_web_team_ga_table

    Range(TabNames.search_output, 'K2').vertical.offset(orders_offset, 0).value = \
        'Total Orders', total_prepaid_orders, total_postpaid_orders, total_orders

    Range(TabNames.search_output, 'K2').vertical.offset(orders_offset + 1, 0).value = \
        'Pre vs. Post Proxy', prepaid_order_percent, postpaid_order_percent, '100%'

    Range(TabNames.search_output, 'K2').vertical.offset(orders_offset + 3, 0).value = \
        'Web Team Orders', total_orders

    Range(TabNames.search_output, 'R2').vertical.offset(gas_offset, 0).value = \
        'Total GAs', total_prepaid_gas, total_postpaid_gas, total_gas

    Range(TabNames.search_output, 'R2').vertical.offset(gas_offset + 1, 0).value = \
        'Pre vs. Post Proxy', prepaid_ga_percent, postpaid_ga_percent, '100%'

    Range(TabNames.search_output, 'R2').vertical.offset(gas_offset + 3, 0).value = \
        'Web Team eGAs', total_gas

    engine_web_team_orders_table['Percent Total'] = engine_web_team_orders_table['Orders'] / \
                                                    engine_web_team_orders_table['Orders'].sum()

    engine_web_team_orders_table['Prepaid Proxy'] = engine_web_team_orders_table['Orders'] * prepaid_order_percent
    engine_web_team_orders_table['Postpaid Proxy'] = engine_web_team_orders_table['Orders'] * postpaid_order_percent

    engine_web_team_orders_table['Site'] = engine_web_team_orders_table['Web Team'].str.replace('DART Search : ', '')
    engine_web_team_orders_table['Site'] = engine_web_team_orders_table['Site'].str.replace(' Web', '')
    engine_web_team_orders_table['Site'] = \
        engine_web_team_orders_table['Site'].str.replace('DART Search: Whistleout', 'Whistleout')

    Range(TabNames.search_output, 'K2', index=False).vertical.offset(orders_offset + 5, 0).value = \
        engine_web_team_orders_table[['Site', 'Percent Total', 'Prepaid Proxy', 'Postpaid Proxy', 'Orders']]

    engine_web_team_ga_table['Percent Total'] = engine_web_team_ga_table['Total GAs'] / \
                                                engine_web_team_ga_table['Total GAs'].sum()

    engine_web_team_ga_table['Prepaid Proxy'] = engine_web_team_ga_table['Total GAs'] * prepaid_ga_percent
    engine_web_team_ga_table['Postpaid Proxy'] = engine_web_team_ga_table['Total GAs'] * postpaid_ga_percent

    engine_web_team_ga_table['Site'] = engine_web_team_ga_table['Web Team'].str.replace('DART Search : ', '')
    engine_web_team_ga_table['Site'] = engine_web_team_ga_table['Site'].str.replace(' Web', '')
    engine_web_team_ga_table['Site'] = \
        engine_web_team_ga_table['Site'].str.replace('DART Search: Whistleout', 'Whistleout')

    Range(TabNames.search_output, 'R2', index=False).vertical.offset(orders_offset + 5, 0).value = \
        engine_web_team_ga_table[['Site', 'Percent Total', 'Prepaid Proxy', 'Postpaid Proxy', 'Total GAs']]


def generate_search_cfv_report():
    wb = Workbook.caller()
    wb.save()

    cfv = search_cfv_report(wb.fullname)

    search_cfv_outputs(cfv)

    Sheet('data').clear_contents()

    unneeded_cols = ['Average CPC', 'Avg CPC', 'CPA', 'CTR', 'CTR (%)', 'Conversion rate (%)', 'Creative Pixel Size',
                     'Transaction Count', 'Creative Type', 'Creative Groups 2', 'Conversions', 'Ad extension type id',
                     'Ad', 'Ad extension ID']

    for i in unneeded_cols:
        if i in list(cfv.columns):
            cfv.drop(i, axis=1, inplace=True)

    datafunc.chunk_df(cfv, 'data', 'A1')

