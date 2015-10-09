import re
from xlwings import Workbook, Range, Sheet
import pandas as pd
from dfa import *
from utility import *
from datafeeds import *
from campaign_reports import *

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

def weekly_reporting():

    wb = Workbook.caller()

    wb.save()

    # Before function is ran, VBA code will create the necessary tabs in order to process correctly. See the
    # documentation in the VBA modules for more information.
    # Workbook needs to be saved in order to load the data into pandas properly
    # Load the Site Activity and Custom Floodlight Variable data into pandas as DataFrames

    sheet = Range('Action_Reference', 'AG1').value

    cfv2 = pd.read_excel(sheet, 'CFV_Temp', index_col=None)
    sa = pd.read_excel(sheet, 'SA_Temp', index_col=None)

    sa_creative = sa[['Placement', 'Creative Field 1']]
    sa_creative.drop_duplicates(subset = 'Placement', inplace = True)

    cfv = pd.merge(cfv2, sa_creative, how = 'left', on = 'Placement')

    cfv = cfv_report.clean_cfv(cfv)

    data = sa.append(cfv)

    data = clickthroughs.strip_clickthroughs(data)

    data = floodlights.run_action_floodlight_tags(data)
    data = categorization.categorize_report(data)

    data = floodlights.f_tags(data)

    data = report_columns.additional(data)

    sa_columns = list(sa.columns)
    tag_columns = sa_columns[sa_columns.index('DBM Cost USD') + 1:]

    columns = report_columns.order() + tag_columns

    data = data[columns]
    data.fillna(0, inplace=True)

    if Range('data', 'A1').value is None:
        chunk_df(data, 'data', 'A1', 5000)

    # If data is already present in the tab, the two data sets are merged together and then copied into the data tab.

    else:

        past_data = pd.read_excel(Range('Action_Reference', 'AG1').value, 'data', index_col=None)
        appended_data = past_data.append(data)
        appended_data = appended_data[columns]
        appended_data.fillna(0, inplace=True)
        Sheet('data').clear()
        chunk_df(appended_data, 'data', 'A1', 5000)

    qa.placement_qa(data)

    ddr_devices.top_15_devices(cfv2)

def data_compression():

    compress.compress_data()

def data_split():

    split_data.split()

def data_merge():

    merge.merge_data()

def ebay_costfeed():

    costfeed()