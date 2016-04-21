import main
import pandas as pd
import numpy as np
import arrow
from xlwings import Application, Workbook, Range


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


def dr_placement_types():
    cd = '|'.join(list(['C/D', 'C Pages', 'D Pages']))
    t2t = '|'.join(list(['Tablet to Tablet']))
    fbx = '|'.join(list(['FBX ']))
    search = '|'.join(list(['Search']))
    pros = '|'.join(list(['Prospecting']))
    aal = '|'.join(list(['Add']))
    tap_att = '|'.join(list(['Tap-to-Call (AT&T)']))
    tap_other = '|'.join(list(['Tap-to-Call (Other)']))
    tap_verizon = '|'.join(list(['Tap-to-Call (Verizon)']))
    tap_sprint = '|'.join(list(['Tap-to-Call (Sprint)']))

    return (cd, t2t, fbx, search, pros, aal, tap_att, tap_other, tap_sprint, tap_verizon)


def transform_dr_forecasts(dr):
    dr = dr[dr['Campaign'] == 'DR']
    ddr = drop_columns(dr)

    cd, t2t, fbx, search, pros, aal, tap_att, tap_other, tap_sprint, tap_verizon = dr_placement_types()

    pacing_dr = ddr.rename(columns={'Placement Messaging Type': 'Type', 'NTC Media Cost': 'Spend', 'Total GAs': 'GAs'})

    pacing_dr['Type.Agg'] = np.where(
        pacing_dr['Type'].str.contains(fbx) == True, 'FBX.Remessaging',
        np.where(pacing_dr['Type'].str.contains(t2t) == True, 'Tablet-to-Tablet',
                 np.where(pacing_dr['Type'].str.contains(cd) == True, 'CD.Remessaging',
                          np.where(pacing_dr['Type'].str.contains(search) == True, 'Search.Remessaging',
                                   np.where(pacing_dr['Type'].str.contains(pros) == True, 'Prospecting',
                                            'Add-A-Line')))))

    pacing_dr = pacing_dr[(pacing_dr['Date'] >= main.quarter_start()) &
                          (pacing_dr['Site'].str.contains(dr_sites()) == True)]

    pacing_dr = pacing_dr.groupby(['Site', 'Type.Agg', 'Type', 'Date'])

    pacing_dr = pd.DataFrame(pacing_dr.sum()).reset_index()

    pacing_dr['Site_Type'] = pacing_dr['Site'] + '.' + pacing_dr['Type.Agg']

    pacing_pivoted = pd.pivot_table(pacing_dr, values=['Spend', 'GAs'], index=['Date'], columns=['Site_Type'],
                               aggfunc=np.sum)
    pacing_pivoted.fillna(0, inplace=True)

    spend = pacing_pivoted['Spend']
    gas = pacing_pivoted['GAs']

    spend.columns = pd.Series(spend.columns).astype(str) + '.Spend'
    gas.columns = pd.Series(gas.columns).astype(str) + '.GAs'

    merged = pd.merge(spend, gas, how= 'left', right_index= True, left_index= True)
    merged.reset_index(inplace=True)

    merged['Weekday'] = merged['Date'].apply(lambda x: arrow.get(x).format('dddd'))

    return merged


def q4_pacing():
    a = '2015-09-28'
    b = '2015-10-04'
    c = '2015-10-05'
    d = '2015-12-27'
    e = '2015-12-28'
    f = '2015-12-31'

    date_rng = pd.date_range(main.quarter_start(), periods=92, freq='D')
    date_df = pd.DataFrame(date_rng, columns=['Date'])

    pacing_wb = Workbook(main.dr_pacing_path())
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
