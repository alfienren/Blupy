import re

import pandas as pd
from xlwings import Range, Sheet

from reporting.constants import TabNames


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


def read_site_activity_report(path, adv='tmo'):
    sa = pd.read_excel(path, TabNames.site_activity, index_col=None)
    if 'DBM Cost USD' in list(sa.columns):
        sa.rename(columns={'DBM Cost USD':'DBM Cost (USD)'}, inplace=True)

    if adv == 'tmo':
        sa_creative = sa[['Placement', 'Creative Field 1']]
        sa_creative = sa_creative.drop_duplicates(subset = 'Placement')

        return sa, sa_creative

    else:
        return sa


def read_cfv_report(path):
    cfv = pd.read_excel(path, TabNames.floodlight_variable, index_col=None)

    return cfv


def merge_past_data(data, columns, path):
    if Range('data', 'A1').value is None:
        chunk_df(data, 'data', 'A1')

    # If data is already present in the tab, the two data sets are merged together and then copied into the data tab.

    else:
        past_data = pd.read_excel(path, 'data', index_col=None)
        appended_data = past_data.append(data)
        appended_data = appended_data[columns]
        appended_data.fillna(0, inplace=True)
        Sheet('data').clear_contents()
        chunk_df(appended_data, 'data', 'A1')

