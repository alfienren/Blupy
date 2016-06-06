import datetime

import pandas as pd
from win32com import client as win32
from xlwings import Range

from reporting import categorization
from reporting.categorization import qquarter
from reporting.constants import DrPerformance


def generate_publisher_emails(pubs_combined, contacts, br):
    week_end = Range(DrPerformance.pub_performance_sheet, 'C7').value + datetime.timedelta(days=6)
    week_end = week_end.strftime('%m/%d').lstrip('0').replace('0', '')
    week_start = Range(DrPerformance.pub_performance_sheet, 'C7').value.strftime('%m/%d').lstrip('0').replace('0', '')

    br['Traffic Yield'] = br['Traffic Yield'].astype(float) * 100
    br['Traffic Yield'] = br['Traffic Yield'].map('{0:.2f}%'.format)

    headerstyle = '<p style = "font-family: Calibri; font-size: 11pt; font-weight: bold; text-decoration: underline;">'
    bodystyle = '<p style = "font-family: Calibri; font-size: 11pt;">'
    boldstyle = '<p style = "font-family: Calibri; font-size: 11pt; font-weight: bold;">'

    merged = pd.merge(pubs_combined, contacts, how= 'left', on= 'Site')

    pub_list =  list(merged['Site'].unique())
    pub_emails = list(merged['cc_emails'].unique())
    contact_emails = list(merged['Contact Email'].unique())

    outlook = win32.Dispatch('Outlook.Application')

    for i in range(0, len(pub_list)):
        if pub_list[i] == 'ASG' or pub_list[i] == 'AOD':
            greeting = 'Hi Andrew and Team,'
            flight_start = '4/18'
        elif pub_list[i] == 'Amazon':
            greeting = 'Hi Kate and Team,'
            flight_start = '4/18'
        elif pub_list[i] == 'eBay':
            greeting = 'Hi Katie,'
            flight_start = '4/18'
        elif pub_list[i] == 'Magnetic':
            greeting = 'Hi Melissa,'
            flight_start = '4/18'
        elif pub_list[i] == 'Yahoo!':
            greeting = 'Hi Krystal,'
            flight_start = '4/18'
        elif pub_list[i] == 'Bazaar Voice':
            greeting = 'Hi Alexa and Todd,'
            flight_start = '4/18'
        elif pub_list[i] == 'Drawbridge':
            greeting = 'Hi Stephanie and Mani,'
            flight_start = '4/18'
        else:
            greeting = 'Hello,'
            flight_start = '4/18'

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
        elif pub_list[i] == 'TripleLift':
            br_performance = headerstyle + \
                'Brand Remessaging</p>' + \
                bodystyle + \
                'Traffic Yield - ' + str(br['Traffic Yield'][2]) + \
                '<br>' + \
                'Traffic Actions - ' + str(int(br['Traffic Actions'][2])) + \
                '</p>' + \
                boldstyle + \
                'Brand Remessaging Performance Chart ' + flight_start + ' - ' + week_end + \
                '</p><br><br><br>'
        else:
            br_performance = ''

        mail = outlook.CreateItem(0)

        df = merged[merged['Site'] == pub_list[i]]

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


def contacts():
    contacts_df = pd.DataFrame(Range('Publisher_Contacts', 'A1').table.value,
                            columns= Range('Publisher_Contacts', 'A1').horizontal.value)
    contacts_df.drop(0, inplace= True)

    return contacts_df


def aggregated():
    pubs_combined = pd.DataFrame(Range(DrPerformance.pub_performance_sheet,
                                       DrPerformance.aggregate_column + str(DrPerformance.row_num)).table.value,
                                 columns=Range(DrPerformance.pub_performance_sheet,
                                               DrPerformance.aggregate_column + str(
                                                   DrPerformance.row_num)).horizontal.value)
    pubs_combined.drop(0, inplace=True)

    return pubs_combined


def brand_remessaging():
    br = pd.DataFrame(Range(DrPerformance.pub_performance_sheet, DrPerformance.br_column + str(
        DrPerformance.row_num)).table.value, columns=Range(DrPerformance.pub_performance_sheet,
                                                           DrPerformance.br_column + str(
                                                               DrPerformance.row_num)).horizontal.value)
    br.drop(0, inplace=True)

    return br


def site_tactic():
    pub_performance = pd.DataFrame(Range(DrPerformance.pub_performance_sheet, DrPerformance.site_tactic_column + str(
        DrPerformance.row_num)).table.value, columns=Range(DrPerformance.pub_performance_sheet,
                                                           DrPerformance.site_tactic_column + str(
                                                               DrPerformance.row_num)).horizontal.value)
    pub_performance.drop(0, inplace=True)

    pub_performance['CPO'] = pub_performance['Spend'].astype(float) / pub_performance['Orders'].astype(float)

    pub_performance.rename(columns= {'Orders':'Total ' + categorization.qquarter() + ' Orders',
                                     'Spend':'Total ' + categorization.qquarter() + ' Spend',
                                     'Placement Messaging Type':'Tactic'}, inplace= True)

    return pub_performance


def tables_for_emails(pub_performance):
    Range(DrPerformance.pub_performance_sheet, 'A30:A1000').clear_contents()

    site_list = list(pub_performance['Site'].unique())

    start_column = 'B'
    site_column = 'A'

    cell_number = Range(DrPerformance.pub_performance_sheet, 'A13').vertical.last_cell.row + 3

    for i in site_list:
        df = pub_performance[pub_performance['Site'] == i]
        df = df[['Tactic', 'Total ' + categorization.qquarter() + ' Orders',
                 'Total ' + categorization.qquarter() + ' Spend', 'CPO', 'Q2 CPO Goal']]
        Range(DrPerformance.pub_performance_sheet, start_column + str(cell_number), index = False).value = df
        Range(DrPerformance.pub_performance_sheet, site_column + str(cell_number), index = False).value = i
        cell_number += 7