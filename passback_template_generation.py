import re
import time
import datetime

from xlwings import Workbook, Range, Sheet
import pandas as pd
import numpy as np
import xlsxwriter
import win32com.client
import os


def passback_generation():

    wb = Workbook.caller()

    placements = pd.DataFrame(Range('passback_placements', 'A1').table.value,
                              columns=Range('passback_placements', 'A1').horizontal.value)

    placements.drop(0, inplace=True)

    site_lookup = pd.DataFrame(Range('site_lookup', 'A1').table.value,
                               columns=Range('site_lookup', 'A1').horizontal.value)

    site_lookup.drop(0, inplace=True)

    ref = pd.merge(placements, site_lookup, how='left', left_on='Site', right_on='DFA Site')
    ref.drop(['Site', 'DFA Site'], axis=1, inplace=True)
    ref.rename(columns={'Site2': 'Site'}, inplace=True)

    ref['Spend'] = np.nan
    ref['Impressions'] = np.nan
    ref['Clicks'] = np.nan
    ref['Video Plays'] = np.nan
    ref['100% Video Completes'] = np.nan

    start = Range('passback_placements', 'F1').value
    end = Range('passback_placements', 'F2').value

    delta = end - start

    days = []
    for i in range(delta.days + 1):
        days.append(datetime.datetime.strftime(start + datetime.timedelta(days=i), '%x'))

    days = pd.DataFrame(days, columns=['Day'])
    dayarray = np.array(days)

    template = pd.DataFrame(np.tile(ref, (len(days['Day']), 1)), columns = ref.columns).join(pd.DataFrame(
                            {'Date': np.repeat(dayarray, len(ref))}))

    template.drop_duplicates(inplace = True)

    columns = ['Date', 'Campaign', 'Site', 'Placement', 'Spend', 'Impressions', 'Clicks', 'Video Plays', '100% Video Completes']

    template = template[columns]

    template.sort('Site', axis = 0, inplace = True)

    sites = template.groupby('Site')

    for i in sites.groups:
        Sheet.add()
        Range('A1', index=False).value = sites.get_group(i)







