import pandas as pd
import numpy as np
import arrow
from reporting import categorization, datafunc
from xlwings import Range, Workbook


def open_planned_media_report():
    plan_sheet = Range('Action_Reference', 'AE1').value
    planned = pd.read_csv(plan_sheet)

    return planned


def site_pacing_report():
    wb = Workbook.caller()
    path = wb.fullname

    planned = open_planned_media_report()

    planned = categorization.sites(planned)

    actual = pd.read_excel(path, 'data', parse_cols='B:AD', index_col=None)

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

