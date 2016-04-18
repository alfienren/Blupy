import main
import data_transform
from xlwings import Workbook, Range
import pandas as pd
import numpy as np
import datetime
import win32com.client as win32


def quarter_start():
    start = '1/1'

    return start


def placement_type_names():
    #cd = '|'.join(list(['C/D Remessaging']))
    cp = '|'.join(list(['C Pages']))
    dp = '|'.join(list(['D Pages']))
    t2t = '|'.join(list(['Tablet to Tablet C&D Remessaging']))
    fbx = '|'.join(list(['FBX Remessaging']))
    search = '|'.join(list(['Search Remessaging']))
    pros = '|'.join(list(['Prospecting']))
    aal = '|'.join(list(['Add-A-Line']))
    tap_att = '|'.join(list(['Tap-to-Call (AT&T)']))
    tap_other = '|'.join(list(['Tap-to-Call (Other)']))
    tap_verizon = '|'.join(list(['Tap-to-Call (Verizon)']))
    tap_sprint = '|'.join(list(['Tap-to-Call (Sprint)']))

    return (cp, dp, t2t, fbx, search, pros, aal, tap_att, tap_other, tap_verizon, tap_sprint)


def goals(data):
    cp, dp, t2t, fbx, search, pros, aal, tap_att, tap_other, tap_verizon, tap_sprint = placement_type_names()
    types = list([cp, dp, t2t, fbx, search, pros, aal, tap_att, tap_other, tap_verizon, tap_sprint])

    goal = pd.DataFrame({'Q1 CPGA Goal':
                             [227.00, 227.00, 450.00, 290.00, 350.00, 725.00, 213.00, 500.00, 500.00, 500.00, 500.00],
                         'Tactic':
                             types})
    merged = pd.merge(data, goal, how= 'left', on= 'Tactic')

    return merged


def week_of(dr):
    week = 'Spend Week of ' + str(dr['Week'].max().strftime('%m/%d/%Y').lstrip('0').replace(' 0', ' '))

    return week

def performance_sheet():
    pub_performance_sheet = 'Publisher Performance'

    return pub_performance_sheet


def site_tactic_table():
    row_number = 13

    return row_number


def brand_remessaging_table():
    row_number = 51

    return row_number


def pub_aggregate_table():
    row_number = 37

    return row_number


def publishers(dr):
    cd, t2t, fbx, search, pros, aal, tap_att, tap_other, tap_sprint, tap_verizon = data_transform.dr_placement_types()
    week = week_of(dr)

    # Publisher Performance
    pub_dr = dr[(dr['Campaign'] == 'DR') & (dr['Date'] >= main.quarter_start())]
    pub_dr = pub_dr.groupby(['Site', 'Placement Messaging Type', 'Week', 'Date'])
    pub_dr = pd.DataFrame(pub_dr.sum()).reset_index()
    #pub_dr = pub_dr[(pub_dr['NTC Media Cost'] != 0)]

    pub_dr['Tactic'] = np.where(pub_dr['Placement Messaging Type'].str.contains(fbx) == True,
                                'FBX Remessaging', pub_dr['Placement Messaging Type'])

    # Quarter
    q_dr = pub_dr[pub_dr['Date'] >= main.quarter_start()]
    q_dr = q_dr.groupby(['Site', 'Tactic'])
    q_dr = pd.DataFrame(q_dr.sum()).reset_index()

    q_dr = goals(q_dr)

    # Last Week
    last_week = pub_dr[pub_dr['Week'] == pub_dr['Week'].max()]
    last_week = last_week.groupby(['Site', 'Tactic'])
    last_week = pd.DataFrame(last_week.sum()).reset_index()

    last_week.rename(columns={'NET Media Cost': week}, inplace= True)

    # Publishers Overall
    sites = q_dr.groupby('Site')
    sites = pd.DataFrame(sites.sum()).reset_index()

    sites['CPGA'] = sites['NET Media Cost'] / sites['Total GAs']

    # Brand Remessaging

    br = dr[dr['Campaign'] == 'Brand Remessaging']
    br_quarter = br[br['Date'] >= main.quarter_start()]

    br_quarter = br_quarter.groupby('Site')
    br_quarter = pd.DataFrame(br_quarter.sum().reset_index())
    br_quarter['Traffic Yield'] = br_quarter['Total Traffic Actions'].astype(float) / \
                                  br_quarter['Impressions'].astype(float)

    pacing_wb = Workbook(main.dr_pacing_path())
    pacing_wb.set_current()

    Range(performance_sheet(), 'A' + str(site_tactic_table()), index=False, header=False).value = q_dr[
        ['Site', 'Tactic', 'Total GAs']]
    Range(performance_sheet(), 'E' + str(site_tactic_table()), index= False, header= False).value = \
        q_dr['NET Media Cost']

    Range(performance_sheet(), 'D' + str(site_tactic_table()), index= False, header= False).value = last_week[week]
    Range(performance_sheet(), 'D' + str(site_tactic_table() - 1)).value = week

    Range(performance_sheet(), 'A' + str(pub_aggregate_table() + 1), index= False, header= False).value = \
        sites[['Site', 'Total GAs', 'CPGA']]

    Range(performance_sheet(), 'C7').value = \
        pub_dr['Week'].max().strftime('%m/%d/%Y').lstrip('0').replace(' 0', ' ')

    Range(performance_sheet(), 'B' + str(brand_remessaging_table()), index= False, header= False).value = \
        br_quarter[['Traffic Yield', 'Impressions']]

    Range(performance_sheet(), 'E' + str(brand_remessaging_table()), index= False, header= False).value = \
        br_quarter['Total Traffic Actions']

    Range(performance_sheet(), 'F' + str(brand_remessaging_table()), index= False, header= False).value = \
        br_quarter['NET Media Cost']

    pacing_wb.save()
    pacing_wb.close()


def publisher_contact_info():
    contacts = pd.DataFrame(Range('Publisher_Contacts', 'A1').table.value,
                            columns= Range('Publisher_Contacts', 'A1').horizontal.value)
    contacts.drop(0, inplace= True)

    return contacts


def publisher_overall_data():
    pubs_combined = pd.DataFrame(Range(performance_sheet(), 'A' + str(pub_aggregate_table())).table.value,
                                 columns= Range(performance_sheet(), 'A' + str(pub_aggregate_table())).horizontal.value)
    pubs_combined.drop(0, inplace= True)

    return pubs_combined


def brand_remessaging():
    br = pd.DataFrame(Range(performance_sheet(), 'A' + str(brand_remessaging_table() - 1)
                            + ':E' + str(brand_remessaging_table() - 1)).vertical.value,
                      columns= Range(performance_sheet(), 'A' + str(brand_remessaging_table() - 1)
                            + ':E' + str(brand_remessaging_table() - 1)).value)
    br.drop(0, inplace=True)

    return br


def publisher_tactic_data():
    pub_performance = pd.DataFrame(Range(performance_sheet(), 'A' + str(site_tactic_table() - 1)
                                         + ':E' + str(site_tactic_table() - 1)).vertical.value,
                                   columns= Range(performance_sheet(), 'A' + str(site_tactic_table() - 1)
                                         + ':E' + str(site_tactic_table() - 1)).value)
    pub_performance.drop(0, inplace= True)

    pub_performance['CPGA'] = pub_performance['Spend'].astype(float) / pub_performance['Total GAs'].astype(float)

    pub_performance.rename(columns= {'Total GAs':'Total ' + main.qquarter() + ' GAs',
                                     'Spend':'Total ' + main.qquarter() + ' Spend',
                                     'Placement Messaging Type':'Tactic'}, inplace= True)

    pub_performance = goals(pub_performance)

    return pub_performance


def generate_publisher_tables(pub_performance):
    Range(performance_sheet(), 'A' + str(brand_remessaging_table() + 2) + ':A1000').clear_contents()

    site_list = list(pub_performance['Site'].unique())
    week = Range(performance_sheet(), 'D' + str(site_tactic_table() - 1)).value

    start_column = 'B'
    site_column = 'A'
    cell_number = brand_remessaging_table() + 3

    for i in site_list:
        df = pub_performance[pub_performance['Site'] == i]
        df = df[['Tactic', 'Total ' + main.qquarter() + ' GAs', week,
                 'Total ' + main.qquarter() + ' Spend', 'CPGA', main.qquarter() + ' CPGA Goal']]
        Range(performance_sheet(), start_column + str(cell_number), index = False).value = df
        Range(performance_sheet(), site_column + str(cell_number), index = False).value = i
        cell_number += 7


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
                'Traffic Actions - ' + str(int(br['Traffic Actions'])) + \
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
        mail.subject = main.qquarter() + ' 2016 DDR Performance Update: ' + \
                       week_start + '-' + week_end + ' - ' + str(pub_list[i])
        mail.HTMLBody = '<body>' + \
            bodystyle + \
            greeting + \
            '<br><br>' + \
            'Below you will find your campaign performance breakout for the beginning of ' + \
                        main.qquarter() + ' through the week ' + \
            'ending ' + week_end + '.</p>' + \
            headerstyle + \
            'DDR Performance ' + flight_start + ' - ' + week_end + ' (All Tactics Combined)</p>' + \
            bodystyle + \
            'CPGA - $' + str(float(df['CPGA']))[:-2] + '<br>' + 'GAs - ' + str(float(df['GAs'])) + \
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
