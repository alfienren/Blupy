from xlwings import Range
import pandas as pd
import numpy as np
import main
import os
from weekly_reporting import categorization


def load_traffic_sheet():
    traffic_sheet = Range('Action_Reference', 'AE1').value
    budgets = pd.read_excel(traffic_sheet, 0)

    budgets = budgets[budgets['Cost structure'].str.contains('Flat rate') == True]

    return budgets


def add_traffic_columns(budgets):
    if not ('Days in Flight' or 'Spend per Day' or 'Units per Day' or 'Placement Group') in budgets.columns:
        budgets['Days in Flight'] = (pd.to_datetime(budgets['End date']) -
                                     pd.to_datetime(budgets['Start date'])) / np.timedelta64(1, 'D')

        budgets['Spend per Day'] = budgets['Cost (USD)'] / budgets['Days in Flight']
        budgets['Units per Day'] = budgets['Units'] / budgets['Days in Flight']

        budgets['Placement Group'] = np.where((budgets['Units'] != 0) | (budgets['Cost (USD)'] != 0),
                                              budgets['Placement'], np.NaN)

    return budgets


def add_new_site_column(budgets):
    if not 'Site' in budgets.columns:
        budgets['Site'] = budgets['Site'].apply(lambda x: x.rsplit(' ', 1)[0])
        budgets['Site'] = budgets['Site'].str.strip()
        budgets.rename(columns={'Site':'Site (DCM)'}, inplace=True)

        budgets = categorization.sites(budgets)

    return budgets


def transform_data_columns(budgets):
    budgets['Units'] = np.where(
        ((budgets['Units'] == 0) | (pd.isnull(budgets['Units']) == True)) &
        ((budgets['Rate (USD)'] == 0) | (pd.isnull(budgets['Rate (USD)']) == True)) &
        ((budgets['Cost (USD)'] == 0) | (pd.isnull(budgets['Cost (USD)']) == True)) &
        ((budgets['Spend per Day'] == 0) | (pd.isnull(budgets['Spend per Day']) == True)) &
        ((budgets['Units per Day'] == 0) | (pd.isnull(budgets['Units per Day']) == True)),
        budgets['Units'].replace(0, np.NaN),
        budgets['Units'])

    budgets['Rate (USD)'] = np.where(
        ((budgets['Units'] == 0) | (pd.isnull(budgets['Units']) == True)) &
        ((budgets['Rate (USD)'] == 0) | (pd.isnull(budgets['Rate (USD)']) == True)) &
        ((budgets['Cost (USD)'] == 0) | (pd.isnull(budgets['Cost (USD)']) == True)) &
        ((budgets['Spend per Day'] == 0) | (pd.isnull(budgets['Spend per Day']) == True)) &
        ((budgets['Units per Day'] == 0) | (pd.isnull(budgets['Units per Day']) == True)),
        budgets['Rate (USD)'].replace(0, np.NaN),
        budgets['Rate (USD)'])

    budgets['Cost (USD)'] = np.where(
        ((budgets['Units'] == 0) | (pd.isnull(budgets['Units']) == True)) &
        ((budgets['Rate (USD)'] == 0) | (pd.isnull(budgets['Rate (USD)']) == True)) &
        ((budgets['Cost (USD)'] == 0) | (pd.isnull(budgets['Cost (USD)']) == True)) &
        ((budgets['Spend per Day'] == 0) | (pd.isnull(budgets['Spend per Day']) == True)) &
        ((budgets['Units per Day'] == 0) | (pd.isnull(budgets['Units per Day']) == True)),
        budgets['Cost (USD)'].replace(0, np.NaN),
        budgets['Cost (USD)'])

    budgets['Spend per Day'] = np.where(
        ((budgets['Units'] == 0) | (pd.isnull(budgets['Units']) == True)) &
        ((budgets['Rate (USD)'] == 0) | (pd.isnull(budgets['Rate (USD)']) == True)) &
        ((budgets['Cost (USD)'] == 0) | (pd.isnull(budgets['Cost (USD)']) == True)) &
        ((budgets['Spend per Day'] == 0) | (pd.isnull(budgets['Spend per Day']) == True)) &
        ((budgets['Units per Day'] == 0) | (pd.isnull(budgets['Units per Day']) == True)),
        budgets['Spend per Day'].replace(0, np.NaN),
        budgets['Spend per Day'])

    budgets['Units per Day'] = np.where(
        ((budgets['Units'] == 0) | (pd.isnull(budgets['Units']) == True)) &
        ((budgets['Rate (USD)'] == 0) | (pd.isnull(budgets['Rate (USD)']) == True)) &
        ((budgets['Cost (USD)'] == 0) | (pd.isnull(budgets['Cost (USD)']) == True)) &
        ((budgets['Spend per Day'] == 0) | (pd.isnull(budgets['Spend per Day']) == True)) &
        ((budgets['Units per Day'] == 0) | (pd.isnull(budgets['Units per Day']) == True)),
        budgets['Units per Day'].replace(0, np.NaN),
        budgets['Units per Day'])

    budgets.fillna(method='ffill', inplace=True)

    return budgets


def by_placement_budgets(budgets):
    budgets = budgets[budgets['Object type'] != 'Placement Group']

    group_uniques = list(budgets['Placement Group'].unique())
    placements = list(budgets['Placement'].unique())

    for i in group_uniques:
        budgets['Units'] = np.where(budgets['Placement Group'] == i,
                                               budgets['Units'] / len(
                                                       budgets[budgets['Placement Group'] == i]),
                                               budgets['Units'])

        budgets['Cost (USD)'] = np.where(budgets['Placement Group'] == i,
                                         budgets['Cost (USD)'] / len(budgets[budgets['Placement Group'] == i]),
                                         budgets['Cost (USD)'])

        budgets['Spend per Day'] = np.where(budgets['Placement Group'] == i,
                                            budgets['Spend per Day'] / len(budgets[budgets['Placement Group'] == i]),
                                            budgets['Spend per Day'])

        budgets['Units per Day'] = np.where(budgets['Placement Group'] == i,
                                                  budgets['Units per Day'] / len(
                                                          budgets[budgets['Placement Group'] == i]),
                                                  budgets['Units per Day'])

    byday_budgets = pd.DataFrame()

    for i in placements:
        df = budgets[budgets['Placement'] == i]

        start = df['Start date'].min()
        rng = pd.DataFrame(pd.date_range(start, periods=df['Days in Flight'].max(), freq='D'), columns=['Flight'])

        merged = pd.merge(df, rng, how='right', left_on='Start date', right_on='Flight')
        merged.fillna(method='ffill', inplace=True)

        byday_budgets = byday_budgets.append(merged)

    return byday_budgets


def merge_traffic_sheets():
    traffic_master = pd.DataFrame()

    folder = Range('Action_Reference', 'AE1').value
    folder_contents = os.listdir(folder)

    for i in folder_contents:
        df = pd.read_csv(folder + '/' + i, index_col=None)
        traffic_master = traffic_master.append(df)

    traffic_master.rename(columns={'Name':'Placement'}, inplace=True)

    traffic_master = add_traffic_columns(traffic_master)
    traffic_master = add_new_site_column(traffic_master)

    column_order = ['Campaign', 'Id', 'Object type', 'Placement', 'Start date', 'End date', 'Days in Flight', 'Type',
                    'Compatibility', 'Dimensions', 'Site (DCM)', 'Site', 'Cost structure', 'Cost (USD)',
                    'Spend per Day', 'Units per Day', 'Rate (USD)', 'Units']

    traffic_master = traffic_master[column_order]

    return traffic_master


def output_flat_rates(byday_budgets):
    main.chunk_df(byday_budgets, 'Flat_Rate', 'A1')


