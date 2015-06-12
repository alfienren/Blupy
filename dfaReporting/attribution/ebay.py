import pandas as pd
import os
import datetime
import ftplib
from win32com.client import Dispatch
from xlwings import Workbook, Range

def costfeed():

    Workbook.caller()

    path = Range('Lookup', 'AA1').value
    server = Range('Lookup', 'S1').value
    user = Range('Lookup', 'S2').value
    password = Range('Lookup', 'S3').value
    confirmation_emails = Range('Lookup', 'U1').vertical.value
    filename = 'EBAY_COST_FEED_' + datetime.date.today().strftime('%Y%m%d') + '.txt'
    output_path = os.path.join(path[:path.rindex('\\')], filename)
    ftp_path = str('STOR' + ' ' + filename)

    data = pd.read_excel(path, 'data', parse_cols= 'B, U, X')
    data.rename(columns={'NTC Media Cost':'Spend'}, inplace= True)
    data['Placement ID'] = data['Placement ID'].astype(int)

    end = data['Date'].max()
    delta = datetime.timedelta(days= 7)
    start = end - delta
    data = data[(data['Date'] >= end - delta)& (data['Date'] <= end)]

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

    mail = 0x0
    outlook = Dispatch('Outlook.Application')

    mail = outlook.CreateItem(mail)
    mail.Subject = 'eBay Cost Feed Uploaded to FTP'
    mail.Body = '''Hello,

The eBay Cost Feed reference table for the period ''' + str(start.date().strftime('%d/%m/%Y')) + ' - ' + \
                str(end.date().strftime('%d/%m/%Y')) + ' has been uploaded to the FTP.'

    mail.To = confirmation_emails
    mail.Send()


