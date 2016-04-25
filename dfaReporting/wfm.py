from xlwings import Range, Sheet, Workbook
from reporting import *
import re
import pandas as pd


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


def generate_reporting():
    wb = Workbook.caller()
    wb.save()

    data = pd.read_excel(wb.fullname, paths.sa_tab_name(), index_col=None)
    data = categorization.sites(data)
    data = categorization.placement_categories(data, adv='wfm')
    data = categorization.date_columns(data)
    data = wfm_columns(data)

    ordered_columns = output_columns(data)

    data = data[ordered_columns]

    if Range('data', 'A1').value is None:
        datafunc.chunk_df(data, 'data', 'A1')

    # If data is already present in the tab, the two data sets are merged together and then copied into the data tab.

    else:
        past_data = pd.read_excel(wb.fullname, 'data', index_col=None)
        appended_data = past_data.append(data)
        appended_data = appended_data[ordered_columns]
        appended_data.fillna(0, inplace=True)
        Sheet('data').clear()
        datafunc.chunk_df(appended_data, 'data', 'A1')

    sheets_to_remove = template.delete_sheets(Sheet.all())
    for i in sheets_to_remove:
        Sheet(i).delete()
