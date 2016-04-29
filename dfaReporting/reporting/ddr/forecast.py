import re

import arrow
import numpy as np
import pandas as pd
from xlwings import Range

import reporting.ddr.performance.common
from reporting import categorization, paths, report_columns
from reporting.ddr.performance import publisher


def generate_forecasts():
    path = paths.path_select()

    r_spend = pd.read_excel(path, 0, index_cols= None)
    r_ga = pd.read_excel(path, 1, index_cols= None)

    r_spend.columns = pd.Series(r_spend.columns).astype(str) + '.Spend'
    r_ga.columns = pd.Series(r_ga.columns).astype(str) + '.GAs'

    dates = pd.DataFrame(pd.date_range(Range('Sheet3', 'AT1').value,
                                       periods=len(r_spend),
                                       freq='D'),
                         columns=['Date'])

    r_spend.reset_index(inplace= True)
    r_ga.reset_index(inplace= True)

    del r_spend['index']
    del r_ga['index']

    a = pd.merge(r_spend, dates, how= 'left', right_index= True, left_index= True)
    merged = pd.merge(a, r_ga, how='left', right_index= True, left_index= True)

    merged.columns = pd.Series(merged.columns).str.replace('.Point.Forecast', '')
    merged.columns = pd.Series(merged.columns).str.replace('.', ' ')
    merged.columns = pd.Series(merged.columns).str.replace('C D', 'CD')
    merged.columns = pd.Series(merged.columns).str.replace('Add A Line', 'Add-A-Line')
    merged.columns = pd.Series(merged.columns).str.replace('Yahoo', 'Yahoo!')

    r_output = merged

    return r_output


def merge_pacing_and_forecasts(r_output):
    raw_pacing = pd.DataFrame(Range('raw_pacing_data', 'A1').table.value,
                              columns= Range('raw_pacing_data', 'A1').horizontal.value)
    raw_pacing.drop(0, inplace= True)

    cols_to_remove = '|'.join(list(['Budget', 'Weekday']))

    raw_pacing = raw_pacing.select(lambda x: not re.search(cols_to_remove, x), axis=1)

    raw_pacing.columns = pd.Series(raw_pacing.columns).str.replace('.Add-A-Line.', ' Add-A-Line ')
    raw_pacing.columns = pd.Series(raw_pacing.columns).str.replace('.Add.A.Line.', ' Add-A-Line ')
    raw_pacing.columns = pd.Series(raw_pacing.columns).str.replace('.Add A Line.', ' Add-A-Line ')
    raw_pacing.columns = pd.Series(raw_pacing.columns).str.replace('.CD.Remessaging.', ' CD Remessaging ')
    raw_pacing.columns = pd.Series(raw_pacing.columns).str.replace('.FBX.Remessaging.', ' FBX Remessaging ')
    raw_pacing.columns = pd.Series(raw_pacing.columns).str.replace('.Tablet-to-Tablet.', ' Tablet to Tablet ')
    raw_pacing.columns = pd.Series(raw_pacing.columns).str.replace('.Search.Remessaging.', ' Search Remessaging ')
    raw_pacing.columns = pd.Series(raw_pacing.columns).str.replace('.Prospecting.', ' Prospecting ')

    raw_pacing['Type'] = 'Actual'
    r_output['Type'] = 'Forecast'

    raw_pacing.dropna(inplace= True)

    forecast_data = raw_pacing.append(r_output)
    forecast_data['Week'] = forecast_data['Date'].apply(lambda x: categorization.mondays(x))

    return forecast_data


def transform_dr_forecasts(dr):
    dr = dr[dr['Campaign'] == 'DR']
    ddr = report_columns.dr_drop_columns(dr)

    cd, t2t, fbx, search, pros, aal, tap_att, tap_other, tap_sprint, tap_verizon = ddr.forecast.dr_placement_types()

    pacing_dr = ddr.rename(columns={'Placement Messaging Type': 'Type', 'NTC Media Cost': 'Spend', 'Total GAs': 'GAs'})

    pacing_dr['Type.Agg'] = np.where(
        pacing_dr['Type'].str.contains(fbx) == True, 'FBX.Remessaging',
        np.where(pacing_dr['Type'].str.contains(t2t) == True, 'Tablet-to-Tablet',
                 np.where(pacing_dr['Type'].str.contains(cd) == True, 'CD.Remessaging',
                          np.where(pacing_dr['Type'].str.contains(search) == True, 'Search.Remessaging',
                                   np.where(pacing_dr['Type'].str.contains(pros) == True, 'Prospecting',
                                            'Add-A-Line')))))

    pacing_dr = pacing_dr[(pacing_dr['Date'] >= reporting.ddr.performance.common.quarter_start_year()) &
                          (pacing_dr['Site'].str.contains(categorization.dr_sites()) == True)]

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
