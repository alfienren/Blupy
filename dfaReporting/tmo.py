from xlwings import Range, Workbook, Sheet, Application
from reporting import *
from reporting.cross_channel import *

import pandas as pd
import numpy as np
import datetime
import os


def generate_weekly_reporting():
    wb = Workbook.caller()
    wb.save()

    # Before function is ran, VBA code will create the necessary tabs in order to process correctly. See the
    # documentation in the VBA modules for more information.
    # Workbook needs to be saved in order to load the data into pandas properly
    # Load the Site Activity and Custom Floodlight Variable data into pandas as DataFrames

    sa, sa_creative = datafunc.read_site_activity_report(wb.fullname, adv='tmo')

    cfv = pd.merge(datafunc.read_cfv_report(wb.fullname), sa_creative, how = 'left', on = 'Placement')

    cfv = custom_variables.custom_variable_columns(cfv)
    cfv = custom_variables.ddr_custom_variables(cfv)

    data = sa.append(cfv)
    data = clickthroughs.strip_clickthroughs(data)

    data = floodlights.a_e_traffic(data, adv='tmo')

    data = categorization.categorize_report(data, adv='tmo')
    data = floodlights.f_tags(data)
    data = report_columns.additional_columns(data, adv='tmo')

    sa_columns = list(sa.columns)
    tag_columns = sa_columns[sa_columns.index('DBM Cost (USD)') + 1:]

    columns = report_columns.order_columns(adv='tmo') + tag_columns

    data = data[columns]
    data.fillna(0, inplace=True)

    datafunc.merge_past_data(data, columns, wb.fullname)

    qa.placement_qa(data)

    Application(wkb=wb).xl_app.Run('Postprocess_Report')


def output_flat_rate_report():
    Workbook.caller()

    qa.flat_rate_corrections()


def pacing_report():
    qa.site_pacing_report()


def cost_feed():
    wb = Workbook.caller()
    path = wb.fullname

    filename = 'EBAY_COST_FEED_' + datetime.date.today().strftime('%Y%m%d') + '.txt'
    output_path = os.path.join(path[:path.rindex('\\')], filename)

    if Range('Action_Reference', 'AC1').value is not None:

        ddrpath = Range('Action_Reference', 'AC1').value
        ddr = pd.read_excel(ddrpath, 'data', parse_cols='X, U, AH')
        ddr['Date'] = pd.to_datetime(ddr['Date'])

        data = pd.read_excel(path, 'data', parse_cols= 'B, AI, AL')
        data = data.append(ddr)

    else:
        data = pd.read_excel(path, 'data', parse_cols= 'B, AI, AL')

    end = data['Date'].max()
    start = end - datetime.timedelta(days=6)
    data = data[(data['Date'] >= start) & (data['Date'] <= end)]

    data.rename(columns={'NTC Media Cost':'Spend'}, inplace= True)
    data.dropna(inplace= True)

    data['Placement ID'] = data['Placement ID'].astype(int)
    data['Date'] = [time.date() for time in data['Date']]

    data = data.groupby(['Placement ID', 'Date'])
    data = pd.DataFrame(data.sum().reset_index())

    data.to_csv(output_path, sep= '|', index= False, encoding= 'utf-8')


def cross_channel_dashboard():
    Workbook.caller()

    search_path, tmo_inputs_path, adscope_path, online_path, social_path, offline_path = Range('Ref', 'B2:B7').value

    dma = sources.dma_lookup()
    competitor = sources.competitors(dma, adscope_path)

    search_dat = sources.search(search_path)

    search_impressions = pd.pivot_table(search_dat, index=['Week'],
                                                values=['Branded Search Impressions'], aggfunc=np.sum)

    data = dashboard.merge_channel_data(sources.online(online_path), sources.offline(dma, offline_path,
                                                                                     search_impressions),
                                        search_dat, sources.social(social_path),
                                        sources.tmo_inputs(tmo_inputs_path))

    data_updates = {'Competitive' : competitor, 'merged_channels' : data[0], 'tmo_volume' : data[1]}

    channels = data[0]
    volume = data[1]
    channels.to_excel('C:/Users/aarschle1/Desktop/cross_channel_test.xlsx')
    volume.to_excel('C:/Users/aarschle1/Desktop/cross_channel_volume.xlsx')

    #data['Week'] = pd.to_datetime(data['Week'])

    for i, j in data_updates.iteritems():
        Sheet(i).clear_contents()
        datafunc.chunk_df(j, i, 'A1')

