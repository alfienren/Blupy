import os
import datetime

import pandas as pd
from xlwings import Range

import main


def cost_feed():
    path = main.report_path()
    filename = 'EBAY_COST_FEED_' + datetime.date.today().strftime('%Y%m%d') + '.txt'
    output_path = os.path.join(path[:path.rindex('\\')], filename)

    if Range('Action_Reference', 'AC1').value is not None:

        ddrpath = Range('Action_Reference', 'AC1').value
        ddr = pd.read_excel(ddrpath, 'Working Data', parse_cols='X, U, AH')
        ddr['Date'] = pd.to_datetime(ddr['Date'])

        data = pd.read_excel(path, 'data', parse_cols= 'B, Y, AB')
        data = data.append(ddr)

    else:

        data = pd.read_excel(path, 'data', parse_cols= 'B, Y, AB')

    end = data['Date'].max()
    start = end - datetime.timedelta(days= 7)
    data = data[(data['Date'] >= start) & (data['Date'] <= end)]

    data.rename(columns={'NTC Media Cost':'Spend'}, inplace= True)
    data.dropna(inplace= True)

    data['Placement ID'] = data['Placement ID'].astype(int)
    data['Date'] = [time.date() for time in data['Date']]

    data = data.groupby(['Placement ID', 'Date'])
    data = pd.DataFrame(data.sum().reset_index())

    data.to_csv(output_path, sep= '|', index= False, encoding= 'utf-8')
