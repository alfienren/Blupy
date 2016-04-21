import arrow
import numpy as np
import pandas as pd
from xlwings import Range, Workbook, Application
import performance
import paths


def drop_columns(dr):
    cols_to_drop = ['Month', 'Tactic', 'Placement Category', 'Message Bucket', 'Message Category',
                    'Message Offer', 'A', 'B', 'C', 'D', 'SLV', 'Awareness Actions', 'Consideration Actions',
                    'PI Traffic', 'PC Traffic', 'NET Media Cost', 'Clicks', 'Prepaid GAs', 'Postpaid GAs',
                    'Prepaid SIMs', 'Postpaid SIMs', 'Prepaid Mobile Internet', 'Postpaid Mobile Internet',
                    'Prepaid phone', 'Postpaid phone', 'AAL', 'New device']

    ddr = dr.drop(cols_to_drop, axis= 1)

    return ddr


def dr_sites():
    sites = '|'.join(list(['AOD', 'ASG', 'Amazon', 'Bazaar Voice', 'eBay', 'Magnetic', 'Yahoo']))

    return sites





def q4_pacing():
    a = '2015-09-28'
    b = '2015-10-04'
    c = '2015-10-05'
    d = '2015-12-27'
    e = '2015-12-28'
    f = '2015-12-31'

    date_rng = pd.date_range(performance.quarter_start_year(), periods=92, freq='D')
    date_df = pd.DataFrame(date_rng, columns=['Date'])

    pacing_wb = Workbook(paths.dr_pacing_path())
    pacing_wb.set_current()

    Application(wkb=pacing_wb).xl_app.Run('Clean_Pacing_Data')

    pacing_data = pd.DataFrame(Range('Q4 DDR Pacing by sub-tactic', 'Z1').table.value,
                               columns=Range('Q4 DDR Pacing by sub-tactic', 'Z1').horizontal.value)

    pacing_data.drop(0, inplace=True)

    Range('Q4 DDR Pacing by sub-tactic', 'Z1').table.clear_contents()
    pacing_wb.close()

    pacing_data.set_index('Week', inplace=True)
    pacing_data = pacing_data.resample('1D', fill_method='pad')

    pacing_data.fillna(method='ffill', inplace= True)
    pacing_data.reset_index(inplace=True)
    pacing_data.rename(columns={'Week': 'Date'}, inplace=True)

    pacing_data = pd.merge(pacing_data, date_df, how='right', left_on='Date', right_on='Date')
    pacing_data.fillna(method='ffill', inplace=True)

    pacing_data.ix[a:b] = pacing_data.ix[a:b] / 4
    pacing_data.ix[c:d] = pacing_data.ix[c:d] / 7
    pacing_data.ix[e:f] = pacing_data.ix[e:f] / 4

    pacing_data.set_index('Date', inplace= True)

    return pacing_data
