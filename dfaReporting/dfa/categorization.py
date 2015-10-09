import datetime
import numpy as np
import pandas as pd
import arrow
from xlwings import Range


def placements(data):

    desktop = '|'.join(list(Range('Lookup', 'H2').vertical.value))
    mobile =  '|'.join(list(Range('Lookup', 'I2').vertical.value))
    video = '|'.join(list(Range('Lookup', 'J2').vertical.value))
    standard = '|'.join(list(Range('Lookup', 'K2').vertical.value))

    data['Platform'] = np.where(data['Placement'].str.contains(desktop) == True, 'Desktop',
                        np.where(data['Placement'].str.contains(mobile) == True, 'Mobile',
                                 'Desktop'))

    data['Creative'] = np.where(data['Placement'].str.contains(video) == True, 'Video',
                        np.where(data['Placement'].str.contains(standard) == True, 'Display',
                                 'Video'))

    data['Category'] = data['Platform'] + ' - ' + data['Creative']

    # Categories are broken down by Platform (mobile/tablet, social, desktop), followed by placement creative (Rich
    # Media, Custom, Remessaging, banners, etc.). Lastly, the placement buy type (dCPM, Flat, CPM, etc.)

    # Words to match on are included in the Lookup tab of the Excel sheet.
    # Example output of categories:
    #   Desktop - Standard - dCPM
    #   Mobile - Custom - Flat

    # mobile = '|'.join(list(Range('Lookup', 'B2:B12').value))
    # tablet = '|'.join(list(Range('Lookup', 'B13:B15').value))
    # social = '|'.join(list(Range('Lookup', 'B16:B18').value))
    #
    # rm = '|'.join(list(Range('Lookup', 'D2:D5').value))
    # custom = '|'.join(list(Range('Lookup', 'D6:D15').value))
    # rem = '|'.join(list(Range('Lookup', 'D16:D28').value))
    # vid = '|'.join(list(Range('Lookup', 'D29:D44').value))
    #
    # dynamic = '|'.join(list(Range('Lookup', 'F2:F3').value))
    # other_buy = '|'.join(list(Range('Lookup', 'F4').value))
    #
    # platform = np.where(data['Placement'].str.contains(mobile) == True, 'Mobile',
    #                     np.where(data['Placement'].str.contains(tablet) == True, 'Tablet',
    #                              np.where(data['Placement'].str.contains(social) == True, 'Social', 'Desktop')))
    #
    # creative = np.where(data['Placement'].str.contains(rm) == True, 'Rich Media',
    #                     np.where(data['Placement'].str.contains(custom) == True, 'Custom',
    #                              np.where(data['Placement'].str.contains(rem) == True, 'Remessaging',
    #                                       np.where(data['Placement'].str.contains(vid) == True, 'Video', 'Standard'))))
    #
    # buy = np.where(data['Placement'].str.contains(dynamic) == True, 'dCPM',
    #                np.where(data['Placement'].str.contains(other_buy), 'Flat', ''))

    # data['Platform'] = platform
    # data['P_Creative'] = creative
    # data['Buy'] = buy
    #
    # data['Category'] = data['Platform'] + ' - ' + data['P_Creative'] + ' - ' + data['Buy']
    #
    # data['Category'] = np.where(data['Category'].str[:3] == ' - ', data['Category'].str[3:], data['Category'])
    # data['Category'] = np.where(data['Category'].str[-3:] == ' - ', data['Category'].str[:-3], data['Category'])
    #
    # data['Platform'] = np.where((data['Platform'].str.contains(mobile) == True) | (data['Platform'].str.contains(tablet) == True), 'Mobile',
    #                               np.where(data['Platform'].str.contains(social) == True, 'Social', 'Desktop'))
    #
    # data['P_Creative'] = np.where(data['P_Creative'].str.contains(vid) == True, 'Video',
    #                               np.where(data['P_Creative'].str.contains(social) == True, np.NaN, 'Display'))
    #
    # data['Category Adjusted'] = data['Platform'] + ' - ' + data['P_Creative']

    return data

def creative(data):

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

    site_ref = pd.DataFrame(Range('Lookup', 'U1').table.value, columns = Range('Lookup', 'U1').horizontal.value)
    site_ref.drop(0, inplace = True)

    data = pd.merge(data, site_ref, left_on= 'Site (DCM)', right_on= 'DFA', how= 'left')
    data.drop('DFA', axis = 1, inplace = True)

    return data

def language(data):

    spanish_campaigns = '|'.join(list(['Spanish', 'Hispanic', 'SL', 'Univision']))

    data['Language'] = np.where(data['Campaign'].str.contains(spanish_campaigns) == True, 'SL', 'EL')

    return data

def mondays(dates):

    monday = dates - datetime.timedelta(days= dates.weekday()) + datetime.timedelta(days= 7, weeks= -1)

    return monday

def date_columns(data):

    # Week
    # Month
    # Quarter

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

def categorize_report(data):

    data = placements(data)
    data = sites(data)
    data = creative(data)
    data = language(data)
    data = date_columns(data)

    return data