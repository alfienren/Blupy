import os
import datetime

import pandas as pd
from xlwings import Workbook, Range


def tmo():

    Workbook.caller()

    path = Range('Action_Reference', 'AA1').value
    server = Range('Action_Reference', 'S1').value
    user = Range('Action_Reference', 'S2').value
    password = Range('Action_Reference', 'S3').value
    filename = 'EBAY_COST_FEED_' + datetime.date.today().strftime('%Y%m%d') + '.txt'
    output_path = os.path.join(path[:path.rindex('\\')], filename)
    ftp_path = str('STOR' + ' ' + filename)

    if Range('Action_Reference', 'AC1').value is not None:

        ddrpath = Range('Action_Reference', 'AC1').value
        ddr = pd.read_excel(ddrpath, 'Working Data', parse_cols='X, U, AH')
        ddr['Date'] = pd.to_datetime(ddr['Date'])

        data = pd.read_excel(path, 'data', parse_cols= 'B, Y, AB')
        data = data.append(ddr)

    else:

        data = pd.read_excel(path, 'data', parse_cols= 'B, Y, AB')

    data.rename(columns={'NTC Media Cost':'Spend'}, inplace= True)
    data.dropna(inplace= True)
    data['Placement ID'] = data['Placement ID'].astype(int)

    data['Date'] = [time.date() for time in data['Date']]

    end = data['Date'].max()
    start = end - datetime.timedelta(days= 7)
    data = data[(data['Date'] >= start) & (data['Date'] <= end)]

    columns = ['Placement ID', 'Date', 'Spend']
    data = data[columns]

    data.to_csv(output_path, sep= '|', index= False, encoding= 'utf-8')

    #ftp = ftplib.FTP(server)
    #ftp.login(user, password)

    #table = open(output_path, 'r')
    #ftp.storlines(ftp_path, table)
    #table.close()
    #ftp.quit()