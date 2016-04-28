import datetime

import pandas as pd
from win32com import client as win32
from xlwings import Range, Workbook, Application

from reporting.ddr.performance.performance import performance_sheet, qquarter
from reporting.ddr.performance.tables import tables_for_emails, site_tactic, aggregated, contacts, brand_remessaging


def generate_publisher_emails(pubs_combined, contacts, br):
    week_end = Range(performance_sheet(), 'C7').value + datetime.timedelta(days=6)
    week_end = week_end.strftime('%m/%d').lstrip('0').replace('0', '')
    week_start = Range(performance_sheet(), 'C7').value.strftime('%m/%d').lstrip('0').replace('0', '')

    br['Traffic Yield'] = br['Traffic Yield'].astype(float) * 100
    br['Traffic Yield'] = br['Traffic Yield'].map('{0:.2f}%'.format)

    headerstyle = '<p style = "font-family: Calibri; font-size: 11pt; font-weight: bold; text-decoration: underline;">'
    bodystyle = '<p style = "font-family: Calibri; font-size: 11pt;">'
    boldstyle = '<p style = "font-family: Calibri; font-size: 11pt; font-weight: bold;">'

    merged = pd.merge(pubs_combined, contacts, how= 'left', on= 'Publisher')

    pub_list =  list(merged['Publisher'].unique())
    pub_emails = list(merged['cc_emails'].unique())
    contact_emails = list(merged['Contact Email'].unique())

    outlook = win32.Dispatch('Outlook.Application')

    for i in range(0, len(pub_list)):
        if pub_list[i] == 'ASG' or pub_list[i] == 'AOD':
            greeting = 'Hi Dan and Team,'
            flight_start = '1/1'
        elif pub_list[i] == 'Amazon':
            greeting = 'Hi Kate and Team,'
            flight_start = '1/1'
        elif pub_list[i] == 'eBay':
            greeting = 'Hi Katie,'
            flight_start = '1/1'
        elif pub_list[i] == 'Magnetic':
            greeting = 'Hi Melissa,'
            flight_start = '1/1'
        elif pub_list[i] == 'Yahoo!':
            greeting = 'Hi Krystal,'
            flight_start = '1/1'
        elif pub_list[i] == 'Bazaar Voice':
            greeting = 'Hi Alexa and Todd,'
            flight_start = '1/1'
        elif pub_list[i] == 'Drawbridge':
            greeting = 'Hi Stephanie and Mani,'
            flight_start = '2/1'
        else:
            greeting = 'Hello,'
            flight_start = '1/1'

        if pub_list[i] == 'ASG' or pub_list[i] == 'AOD':
            br_performance = headerstyle + \
                'Brand Remessaging</p>' + \
                bodystyle + \
                'Traffic Yield - ' + str(br['Traffic Yield'][1]) + \
                '<br>' + \
                'Traffic Actions - ' + str(int(br['Traffic Actions'][1])) + \
                '</p>' + \
                boldstyle + \
                'Brand Remessaging Performance Chart ' + flight_start + ' - ' + week_end + \
                '</p><br><br><br>'
        else:
            br_performance = ''

        mail = outlook.CreateItem(0)

        df = merged[merged['Publisher'] == pub_list[i]]

        mail.To = str(contact_emails[i]).encode('utf-8')
        mail.CC = str(pub_emails[i]).encode('utf-8')
        mail.subject = qquarter() + ' 2016 DDR Performance Update: ' + \
                       week_start + '-' + week_end + ' - ' + str(pub_list[i])
        mail.HTMLBody = '<body>' + \
            bodystyle + \
            greeting + \
            '<br><br>' + \
            'Below you will find your campaign performance breakout for the beginning of ' + \
                        qquarter() + ' through the week ' + \
            'ending ' + week_end + '.</p>' + \
            headerstyle + \
            'DDR Performance ' + flight_start + ' - ' + week_end + ' (All Tactics Combined)</p>' + \
            bodystyle + \
            'CPO - $' + str(float(df['CPO']))[:-2] + '<br>' + 'Orders - ' + str(float(df['Orders'])) + \
            '</p>' + \
            boldstyle + \
            'DDR Performance Chart by Tactic ' + flight_start + ' - ' + week_end + \
            '<br><br><br><br><br><br><br>' + \
            'Optimization Notes</p>' + \
            br_performance + \
            bodystyle + \
            '<br><br><br>' + \
            'Let me know if you have any questions.' + \
            '<br><br>' + \
            'Best,' + \
            '</body>'
        mail.Display()


def emails_to_publisher():
    pacing_wb = Workbook.caller()

    tables_for_emails(site_tactic())

    generate_publisher_emails(aggregated(), contacts(), brand_remessaging())

    Application(wkb=pacing_wb).xl_app.Run('Format_Tables')