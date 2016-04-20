from xlwings import Workbook
from weekly_reporting import *
from reporting_data import *
from paths import *
from campaign_reports.ddr.dashboard import *

import pandas as pd


def generate_weekly_reporting():
    wb = Workbook.caller()
    wb.save()

    # Before function is ran, VBA code will create the necessary tabs in order to process correctly. See the
    # documentation in the VBA modules for more information.
    # Workbook needs to be saved in order to load the data into pandas properly
    # Load the Site Activity and Custom Floodlight Variable data into pandas as DataFrames

    sa, sa_creative = read_site_activity_report(wb.fullname, adv='tmo')

    cfv = pd.merge(read_cfv_report(wb.fullname), sa_creative, how = 'left', on = 'Placement')

    cfv = custom_variables.custom_variable_columns(cfv)
    cfv = custom_variables.ddr_custom_variables(cfv)

    data = sa.append(cfv)
    data = clickthroughs.strip_clickthroughs(data)

    data = custom_variables.format_custom_variable_columns(data)
    data = floodlights.a_e_traffic(data, adv='tmo')

    data = categorization.categorize_report(data, adv='tmo')
    data = floodlights.f_tags(data)
    data = report_columns.additional_columns(data, adv='tmo')

    sa_columns = list(sa.columns)
    tag_columns = sa_columns[sa_columns.index('DBM Cost (USD)') + 1:]

    columns = report_columns.order_columns(adv='tmo') + tag_columns

    data = data[columns]
    data.fillna(0, inplace=True)

    merge_past_data(data, columns)

    qa.placement_qa(data)


def generate_metro_reporting():
    wb = Workbook.caller()
    wb.save()

    sa = read_site_activity_report(wb.fullname, adv='metro')
    cfv = read_cfv_report(wb.fullname)

    data = sa.append(cfv)
    data = floodlights.a_e_traffic(data, adv='metro')
    data = categorization.date_columns(data)
    data = categorization.categorize_report(data, adv='metro')

    data = report_columns.additional_columns(data, adv='metro')

    sa_columns = list(sa.columns)
    tag_columns = sa_columns[sa_columns.index('Clicks') + 1:]

    columns = report_columns.order_columns(adv='metro') + tag_columns

    data = data[columns]
    data = data.fillna(0)

    merge_past_data(data, columns)

    qa.placement_qa(data)


def generate_wfm_reporting():
    wfm.generate_reporting()


def generate_dr_reporting():
    dr.dr_weekly_reporting()


def dr_generate_dashboard_data():
    wb = Workbook.caller()

    save_path = str(dr_pivot_path())
    save_path = save_path[:save_path.rindex('\\')]

    ddr_data = display.raw_pivot()

    ddr_display = display.tableau_campaign_data(ddr_data)
    ddr_search_data = search.merge_search_data()

    tableau_search = search.tableau_search_data(ddr_search_data)

    tableau = ddr_display.append(tableau_search)

    if Range('merged', 'A1').value is None:
        chunk_df(tableau, 'merged', 'A1')

    # If data is already present in the tab, the two data sets are merged together and then copied into the data tab.
    else:
        past_data = pd.read_excel(dr_pacing_path(), 'merged', index_col=None)
        past_data = past_data[past_data['Campaign'] != 'Search']
        appended_data = past_data.append(tableau)
        Sheet('merged').clear()
        chunk_df(appended_data, 'merged', 'A1')

    search.search_data_client(ddr_search_data, save_path)

    wb2 = Workbook()
    Sheet('Sheet1').name = 'DDR Data'

    chunk_df(ddr_data, 'DDR Data', 'A1')

    wb2.save(save_path + '\\' + 'DR_Raw_Data.csv')
    wb2.close()
