import datetime

import numpy as np
import pandas as pd
import arrow
from xlwings import Range


def placement_categories(data, adv='tmo'):
    desktop = '|'.join(list(Range('Lookup', 'A2').vertical.value))
    mobile =  '|'.join(list(Range('Lookup', 'B2').vertical.value))
    video = '|'.join(list(Range('Lookup', 'C2').vertical.value))
    standard = '|'.join(list(Range('Lookup', 'D2').vertical.value))
    tmob = '|'.join(list(['T-Mobile', 'T-Mob']))

    rm = '|'.join(list(Range('Lookup', 'E2').vertical.value))
    custom = '|'.join(list(Range('Lookup', 'F2').vertical.value))
    rem = '|'.join(list(Range('Lookup', 'G2').vertical.value))
    social = '|'.join(list(Range('Lookup', 'H2').vertical.value))

    data['Placement2'] = np.where(data['Placement'].str.contains(tmob) == True,
                                    data['Placement'].str.replace(tmob, ''), data['Placement'])


    data['Platform'] = np.where((data['Placement2'].str.contains(mobile) == True), 'Mobile',
                        np.where(data['Placement2'].str.contains(desktop) == True, 'Desktop',
                                 'Desktop'))

    data['Creative2'] = np.where(data['Placement2'].str.contains(video) == True, 'Video',
                        np.where(data['Placement'].str.contains(standard) == True, 'Display',
                                 'Display'))

    data['Creative3'] = np.where(data['Placement2'].str.contains(rm) == True, 'Rich Media',
                                 np.where(data['Placement2'].str.contains(custom) == True, 'Custom',
                                          np.where(data['Placement2'].str.contains(rem) == True, 'Remessaging',
                                                   np.where(data['Placement2'].str.contains(social) == True, 'Social',
                                                            'Standard'))))

    if adv == 'tmo':
        data['Category'] = data['Platform'] + ' - ' + data['Creative2'] + ' - ' + data['Creative3']
        data['Category_Adjusted'] = data['Platform'] + ' - ' + data['Creative2']

    elif adv == 'metro':
        data['TMO_Category'] = data['Platform'] + ' - ' + data['Creative2'] + ' - ' + data['Creative3']
        data['TMO_Category_Adjusted'] = data['Platform'] + ' - ' + data['Creative2']

        cat = pd.DataFrame(Range('Lookup', 'J1').table.value, columns = Range('Lookup', 'J1').horizontal.value)
        cat.drop(0, inplace = True)
        cat.drop_duplicates(subset='Placement Name', inplace=True)

        data = pd.merge(data, cat, left_on = 'Placement', right_on = 'Placement Name', how = 'left')
        data.drop('Placement Name', axis = 1, inplace =True)

        data.drop_duplicates(inplace=True)

    return data


def creative_categories(data):
    # Message Bucket, Category and Offer
    # 90% of the time, the message bucket, category and offer can be determined from the creative field 1 column. It
    # follows a pattern of Bucket_Category_Offer
    data['Creative Field 1'] = data['Creative Field 1'].str.replace('Creative Type: ', '')
    data['Creative Field 1'] = data['Creative Field 1'].str.replace('(', '')
    data['Creative Field 1'] = data['Creative Field 1'].str.replace(')', '')

    # If Creative Field 1 is equal to (not set), this is either a 1x1 or a placement with logo creative. (not set)
    # fields are therefore set as 'TMO Unique Creative', which is how this has been handled historically.

    # Message Bucket is determined by splitting Creative Field 1 and taking the first word.
    data['Message Bucket'] = data['Creative Field 1'].str.split('_').str.get(0)

    # Message Category is determined by splitting Creative Field 1 and taking the second word.
    data['Message Category'] = data['Creative Field 1'].str.split('_').str.get(1)

    # Message Offer is determined by splitting Creative Field 1 and taking the third word. If the offer is not set,
    # it can sometimes be found in the Creative Groups 2 column. For blanks in the Message Offer column, it will try
    # to pull in the offer from the Creative Groups 2 column.

    data['Creative Bucket'] = data['Creative Field 1'].str.split('_').str.get(2)
    data['Creative Bucket'].fillna(data['Creative Field 1'], inplace=True)

    data['Creative Theme'] = data['Creative Groups 2']

    return data


def sites(data):
    site_ref = pd.DataFrame(Range('Lookup', 'Q1').table.value, columns = Range('Lookup', 'Q1').horizontal.value)
    site_ref.drop(0, inplace=True)
    site_ref.drop_duplicates(subset='DFA', inplace=True)

    data = pd.merge(data, site_ref, left_on= 'Site (DCM)', right_on= 'DFA', how= 'left')
    data.drop('DFA', axis = 1, inplace = True)

    return data


def spanish_keywords():
    spanish_campaigns = '|'.join(list(['Spanish', 'Hispanic', 'SL', 'Latino', 'Univision', 'Telemundo']))

    return spanish_campaigns


def media_plan(data):
    spanish_campaigns = spanish_keywords()

    data['Campaign2'] = np.where(data['Campaign'].str.contains('BidManager') == True, data['Placement'],
                                 data['Campaign'])

    data['Media Plan'] = np.where(data['Campaign2'].str.contains('Brand Remessaging|Brand Rms') == True,
                                  'Brand Remessaging',
                                  np.where(data['Campaign2'].str.contains('DDR') == True, 'DDR',
                                           np.where(data['Campaign2'].str.contains('Forbes') == True,
                                                    'Forbes Sponsorship',
                                                    np.where(data['Campaign2'].str.contains('FEP|ADU|OXYGEN') == True,
                                                             'FEP Upfront',
                                                             np.where(data['Campaign2'].str.contains(
                                                                 'DemandGen|Demand Gen') == True, 'Demand Gen',
                                                                      np.where(data['Campaign2'].str.contains(
                                                                          'Network') == True, 'Network',
                                                                               'Super Bowl'))))))

    data['Media Plan'] = np.where(data['Campaign'].str.contains(spanish_campaigns) == True, 'SL ' + data['Media Plan'],
                                  data['Media Plan'])

    data.drop('Campaign2', axis=1, inplace=True)

    return data


def message_campaign(data):
    message_table = pd.DataFrame(Range('Lookup', 'T1').table.value, columns=Range('Lookup', 'T1').horizontal.value)
    message_table.drop(0, inplace=True)

    message_table.drop_duplicates(inplace=True)

    data_merged = pd.merge(data, message_table, on=['Media Plan', 'Creative Groups 2'], how='left')

    return data_merged


def language(data):
    spanish_campaigns = spanish_keywords()

    data['Language'] = np.where(data['Campaign'].str.contains(spanish_campaigns) == True, 'SL', 'EL')

    return data


def mondays(dates):
    monday = dates - datetime.timedelta(days= dates.weekday()) + datetime.timedelta(days= 7, weeks= -1)

    return monday


def date_columns(data):
    quarters = {
        'January': 'Q1',
        'February': 'Q1',
        'March': 'Q1',
        'April': 'Q2',
        'May': 'Q2',
        'June': 'Q2',
        'July': 'Q3',
        'August': 'Q3',
        'September': 'Q3',
        'October': 'Q4',
        'November': 'Q4',
        'December': 'Q4'
    }

    data['Date2'] = pd.to_datetime(data['Date'])

    data['Month'] = data['Date2'].apply(lambda x: arrow.get(x).format('MMMM'))
    data['Quarter'] = data['Month'].apply(lambda x: quarters[x])
    data['Week'] = data['Date2'].apply(lambda x: mondays(x))
    data.drop('Date2', axis = 1, inplace = True)

    return data


def dr_placement_message_type(data):
    message_type = pd.DataFrame(Range('Lookup', 'K3').table.value, columns = Range('Lookup', 'K3').horizontal.value)
    message_type.drop(0, inplace= True)

    data = pd.merge(data, message_type, left_on= 'Placement', right_on= 'Placement_category', how= 'left')
    data.drop(['Campaign2', 'Placement_category'], axis= 1, inplace= True)

    return data


def dr_tactic(data):
    tactic = pd.DataFrame(Range('Lookup', 'AD1').table.value, columns = Range('Lookup', 'AD1').horizontal.value)
    tactic.drop(0, inplace=True)

    data = pd.merge(data, tactic, on='Placement Messaging Type', how='left')

    return data


def dr_creative_categories(data):
    creative_categories = pd.DataFrame(Range('Lookup', 'Y1').table.value,
                                       columns = Range('Lookup', 'Y1').horizontal.value)
    creative_categories.drop(0, inplace=True)

    data = pd.merge(data, creative_categories, on='Creative Groups 2', how='left')

    return data


def categorize_report(data, adv='tmo'):
    data = sites(data)

    if adv == 'tmo':
        data = placement_categories(data, adv='tmo')
        data = creative_categories(data)
        data = dr_placement_message_type(data)
        data = media_plan(data)
        data = message_campaign(data)
    else:
        data = placement_categories(data, adv='metro')

    data = language(data)
    data = date_columns(data)

    return data
