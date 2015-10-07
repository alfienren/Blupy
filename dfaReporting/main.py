from dfa import *
from utility import *
from datafeeds import *

from xlwings import Workbook, Range, Sheet
import pandas as pd
import numpy as np
import re

def chunk_df(df, sheet, startcell, chunk_size):

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

def dfa_additional_columns(data):

    # The DFA field DBM Cost is more accurate for placements using dynamic bidding. If a placement is not using
    # dynamic bidding, DBM Cost = 0. Therefore, if DBM cost does not equal 0, replace the row's media cost with
    # DBM cost. If DBM Cost = 0, Media Cost stays the same.
    data['Media Cost'] = np.where(data['DBM Cost USD'] != 0, data['DBM Cost USD'], data['Media Cost'])

    # Adjust spend to Net to Client
    data['NTC Media Cost'] = 0

    # DBM Cost column is then removed as it is no longer needed.
    data.drop('DBM Cost USD', 1, inplace=True)

    # Add columns for Video Completions and Views, primarily for compatibility between campaigns that run video and
    # don't run video. Those that don't can keep the columns set to zero, but those that have video can then be
    # adjusted with passback data.
    data['Video Completions'] = 0
    data['Video Views'] = 0

    return data

def order_columns():

    dimensions = ['Week', 'Date', 'Month', 'Quarter', 'Campaign', 'Language', 'Site (DCM)', 'Site', 'Click-through URL',
                  'F Tag', 'Category', 'Message Bucket', 'Message Category', 'Creative Bucket', 'Creative Theme',
                  'Creative Type', 'Creative Groups 1', 'Creative ID', 'Creative', 'Ad', 'Creative Groups 2',
                  'Creative Field 1', 'Placement', 'Placement ID', 'Placement Cost Structure']

    cfv_floodlight_columns = ['OrderNumber (string)',  'Activity','Floodlight Attribution Type',
                              'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)']

    metrics = ['Media Cost', 'NTC Media Cost', 'Impressions', 'Clicks', 'Orders', 'Plans', 'Add-a-Line',
               'Activations', 'Devices', 'Services', 'Accessories',
               'Postpaid Plans', 'Prepaid Plans', 'eGAs', 'Store Locator Visits', 'A Actions', 'B Actions', 'C Actions',
               'D Actions', 'E Actions', 'F Actions', 'Awareness Actions', 'Consideration Actions',
               'Traffic Actions', 'Post-Click Activity', 'Post-Impression Activity', 'Video Views',
               'Video Completions', 'Prepaid GAs', 'Postpaid GAs', 'Prepaid SIMs', 'Postpaid SIMs',
               'Prepaid Mobile Internet', 'Postpaid Mobile Internet', 'Prepaid Phone', 'Postpaid Phone',
               'DDR New Devices', 'DDR Add-a-Line']

    new_columns = dimensions + metrics + cfv_floodlight_columns

    return list(new_columns)

def weekly_reporting():

    wb = Workbook.caller()

    wb.save()

    # Before function is ran, VBA code will create the necessary tabs in order to process correctly. See the
    # documentation in the VBA modules for more information.
    # Workbook needs to be saved in order to load the data into pandas properly
    # Load the Site Activity and Custom Floodlight Variable data into pandas as DataFrames

    cfv = pd.DataFrame(Range('CFV_Temp', 'A1').table.value, columns = Range('CFV_Temp', 'A1').horizontal.value)
    cfv.drop(0, inplace = True)

    sheet = Range('Lookup', 'AA1').value
    sa = pd.read_excel(sheet, 'SA_Temp', index_col=None)

    cfv = cfv_report.run_cfv_macro(cfv)

    data = sa.append(cfv)

    data = clickthroughs.strip_clickthroughs(data)

    data = floodlights.run_action_floodlight_tags(data)
    data = categorization.categorize_report(data)

    data = floodlights.f_tags(data)

    sa_columns = list(sa.columns)
    tag_columns = sa_columns[sa_columns.index('DBM Cost USD') + 1:]

    columns = order_columns() + tag_columns

    data = data[columns]
    data.fillna(0, inplace=True)

    if Range('data', 'A1').value is None:
        chunk_df(data, 'data', 'A1', 5000)

    # If data is already present in the tab, the two data sets are merged together and then copied into the data tab.

    else:

        past_data = pd.read_excel(Range('Lookup', 'AA1').value, 'data', index_col=None)
        appended_data = past_data.append(data)
        appended_data = appended_data[columns]
        appended_data.fillna(0, inplace=True)
        Sheet('data').clear()
        chunk_df(appended_data, 'data', 'A1', 3000)

def data_compression():

    compress.compress_data()

def data_split():

    split_data.split()

def data_merge():

    merge.merge_data()

def ebay_costfeed():

    costfeed()