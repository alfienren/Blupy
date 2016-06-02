import arrow
import numpy as np
import pandas as pd
from xlwings import Sheet, Range, Workbook

import reporting.datafunc
from reporting import categorization, datafunc
from reporting.constants import TabNames

from reporting.datafunc import open_planned_media_report


# Check placements with spend and impressions but 0 clicks. Also check for placements with spend and clicks but
# no impressions.
def placement_qa(data):
    data['Flag'] = np.where((data['Media Cost'] > 50) & (data['Impressions'] < 1000), 'Low Impressions',
                            np.where((data['Media Cost'] > 10) & (data['Impressions'] > 100) & (data['Clicks'] == 0),
                                     'Zero Clicks',
                                     np.where(
                                         (data['Media Cost'] > 0) & (data['Clicks'] > 0) & (data['Impressions'] == 0),
                                         'Zero Impressions', np.nan)))

    data = data[data['Flag'] != 'nan']
    data = data[['Placement', 'Placement ID', 'Week', 'Media Cost', 'Impressions', 'Clicks', 'Flag']]

    Sheet.add('Data_QA_Output', after = 'data')

    Range(TabNames.qa_tab_name, 'A1', index = False).value = data


def flat_rate_corrections():
    planned = reporting.datafunc.open_planned_media_report()
    planned = categorization.sites(planned)

    planned = planned[planned['Placement Cost Structure'].str.contains('Flat Rate') == True]

    if Range('data', 'B1').value == 'Date':
        dates = pd.DataFrame(Range('data', 'B1').vertical.value, columns=['Date'])
        dates.drop(0, inplace=True)
    else:
        dates = pd.DataFrame(Range('data', 'E1').vertical.value, columns=['Date'])
        dates.drop(0, inplace=True)

    max_date = dates['Date'].max()

    planned['Ended'] = np.where(pd.to_datetime(planned['Placement End Date']) <= max_date, 'Ended', '')

    Sheet('Flat_Rate').clear_contents()
    datafunc.chunk_df(planned, 'Flat_Rate', 'A1')


def site_pacing_report():
    wb = Workbook.caller()
    path = wb.fullname

    planned = open_planned_media_report()

    planned = categorization.sites(planned)

    actual = pd.read_excel(path, 'data', parse_cols='B:AJ', index_col=None)

    actual_columns_keep = ['Campaign', 'Site', 'Date', 'Month', 'NTC Media Cost']

    actual = actual[actual_columns_keep]

    start_date = actual['Date'].min().strftime('%m%d%Y')
    end_date = actual['Date'].max().strftime('%m%d%Y')

    output_path = path[:path.rindex('\\')] + '/' + 'Pacing_' + start_date + '-' + end_date + '.xlsx'

    planned['id'] = planned['Month'] + planned['Package/Roadblock']
    planned['id count'] = planned.groupby(['id'])['Placement Total Planned Media Cost'].transform('count')
    planned['planned'] = planned['Placement Total Planned Media Cost'] / planned['id count']
    planned['month count'] = np.round((pd.to_datetime(planned['Placement End Date']) -
                                       pd.to_datetime(planned['Placement Start Date'])) /
                                      np.timedelta64(1, 'M'), decimals=0)

    planned['Monthly Planned'] = np.where(planned['month count'] != 0, planned['planned'] /
                                          planned['month count'], planned['planned'])

    planned['Month'] = pd.to_datetime(planned['Month'])
    planned['Month'] = planned['Month'].apply(lambda x: arrow.get(x).format('MMMM'))

    planned = planned.groupby(['Campaign', 'Site', 'Month'])
    planned = pd.DataFrame(planned.sum()).reset_index()

    actual = actual.groupby(['Campaign', 'Site', 'Month'])
    actual = pd.DataFrame(actual.sum()).reset_index()

    merged = pd.merge(planned, actual, how='left', on=['Campaign', 'Site', 'Month'])

    del merged['Placement Total Planned Media Cost']
    del merged['Planned Media Cost']
    del merged['id count']
    del merged['planned']
    del merged['month count']

    pacing_sheet = Workbook()
    datafunc.chunk_df(merged, 0, 'A1')

    pacing_sheet.save(output_path)
    pacing_sheet.close()

    wb.set_current()