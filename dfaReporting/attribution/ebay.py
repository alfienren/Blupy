import pandas as pd
import os
import datetime
import ftplib
from xlwings import Workbook, Range

def costfeed():

    Workbook.caller()

    path = Range('Lookup', 'AA1').value
    server = Range('Lookup', 'S1').value
    user = Range('Lookup', 'S2').value
    password = Range('Lookup', 'S3').value
    filename = 'EBAY_COST_FEED_' + datetime.date.today().strftime('%Y%m%d') + '.txt'
    output_path = os.path.join(path[:path.rindex('\\')], filename)
    ftp_path = str('STOR' + ' ' + filename)

    if Range('Lookup', 'AC1').value is not None:

        ddrpath = Range('Lookup', 'AC1').value
        ddr = pd.read_excel(ddrpath, 'Working Data', parse_cols='X, U, AH')
        ddr['Date'] = pd.to_datetime(ddr['Date'])

        data = pd.read_excel(path, 'data', parse_cols= 'B, U, X')
        data = data.append(ddr)

    else:

        data = pd.read_excel(path, 'data', parse_cols= 'B, U, X')

    data.rename(columns={'NTC Media Cost':'Spend'}, inplace= True)
    data.dropna(inplace= True)
    data['Placement ID'] = data['Placement ID'].astype(int)

    end = data['Date'].max()
    delta = datetime.timedelta(days= 7)
    start = end - delta
    data = data[(data['Date'] >= start) & (data['Date'] <= end)]

    data['Date'] = [time.date() for time in data['Date']]

    columns = ['Placement ID', 'Date', 'Spend']
    data = data[columns]

    data.to_csv(output_path, sep= '|', index= False, encoding= 'utf-8')

    ftp = ftplib.FTP(server)
    ftp.login(user, password)

    table = open(output_path, 'r')
    ftp.storlines(ftp_path, table)
    table.close()
    ftp.quit()




