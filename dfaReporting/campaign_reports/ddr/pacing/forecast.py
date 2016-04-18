import main
import numpy as np
from xlwings import Range, Sheet, Workbook
import re

import pandas as pd


def generate_forecasts():
    r_spend = pd.read_excel(main.r_output_path(), 0, index_cols= None)
    r_ga = pd.read_excel(main.r_output_path(), 1, index_cols= None)

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
    forecast_data['Week'] = forecast_data['Date'].apply(lambda x: main.monday_week_start(x))

    return forecast_data


def output_forecasts(pacing_data):
    pacing_data['Week'] = pacing_data['Date'].apply(lambda x: main.monday_week_start(x))

    pacing_data = pd.pivot_table(pacing_data, index= ['Site', 'Tactic', 'Metric'],
                          columns= ['Week'], values= 'value', aggfunc= np.sum).reset_index()

    wb = Workbook(main.dr_pacing_path())

    Sheet('forecast_data').clear_contents()
    Range('forecast_data', 'A1', index= False).value = pacing_data

    wb.save()
    wb.close()
