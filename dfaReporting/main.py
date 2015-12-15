import re

from xlwings import Workbook, Range, Sheet
import pandas as pd

from weekly_reporting import *
from datafeeds import *


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


def cfv_tab_name():
    cfv = 'CFV_Temp'

    return cfv


def sa_tab_name():
    sa = 'SA_Temp'

    return sa


def report_path():
    path = Range('Action_Reference', 'AG1').value

    return path


def read_site_activity_report():
    sa = pd.read_excel(report_path(), sa_tab_name(), index_col=None)

    sa_creative = sa[['Placement', 'Creative Field 1']]
    sa_creative.drop_duplicates(subset = 'Placement', inplace = True)

    return (sa, sa_creative)


def read_cfv_report():
    cfv = pd.read_excel(report_path(), cfv_tab_name(), index_col=None)

    return cfv


def generate_weekly_reporting():
    wb = Workbook.caller()

    wb.save()

    # Before function is ran, VBA code will create the necessary tabs in order to process correctly. See the
    # documentation in the VBA modules for more information.
    # Workbook needs to be saved in order to load the data into pandas properly
    # Load the Site Activity and Custom Floodlight Variable data into pandas as DataFrames

    sa, sa_creative = read_site_activity_report()

    cfv = pd.merge(read_cfv_report(), sa_creative, how = 'left', on = 'Placement')

    cfv = custom_variables.custom_variable_columns(cfv)
    cfv = custom_variables.ddr_custom_variables(cfv)

    data = sa.append(cfv)
    data = clickthroughs.strip_clickthroughs(data)

    data = custom_variables.format_custom_variable_columns(data)
    data = floodlights.a_e_traffic(data)

    data = categorization.categorize_report(data)
    data = floodlights.f_tags(data)
    data = report_columns.additional_columns(data)

    sa_columns = list(sa.columns)
    tag_columns = sa_columns[sa_columns.index('DBM Cost USD') + 1:]

    columns = report_columns.order_columns() + tag_columns

    data = data[columns]
    data.fillna(0, inplace=True)

    if Range('data', 'A1').value is None:
        chunk_df(data, 'data', 'A1')

    # If data is already present in the tab, the two data sets are merged together and then copied into the data tab.

    else:
        past_data = pd.read_excel(report_path(), 'data', index_col=None)
        appended_data = past_data.append(data)
        appended_data = appended_data[columns]
        appended_data.fillna(0, inplace=True)
        Sheet('data').clear()
        chunk_df(appended_data, 'data', 'A1')

    qa.placement_qa(data)

    #ddr_devices.top_15_devices(cfv2)


def tmo_costfeed():
    Workbook.caller()

    tmo.cost_feed()


