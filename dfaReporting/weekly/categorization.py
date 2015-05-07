import numpy as np
from xlwings import Range

def categories(data):

    # Categories are broken down by Platform (mobile/tablet, social, desktop), followed by placement creative (Rich
    # Media, Custom, Remessaging, banners, etc.). Lastly, the placement buy type (dCPM, Flat, CPM, etc.)

    # Words to match on are included in the Lookup tab of the Excel sheet.
    # Example output of categories:
    #   Desktop - Standard - dCPM
    #   Mobile - Custom - Flat
    mobile = '|'.join(list(Range('Lookup', 'B2:B6').value))
    tablet = '|'.join(list(Range('Lookup', 'B7:B9').value))
    social = '|'.join(list(Range('Lookup', 'B10:B12').value))

    rm = '|'.join(list(Range('Lookup', 'D2:D5').value))
    custom = '|'.join(list(Range('Lookup', 'D6:D15').value))
    rem = '|'.join(list(Range('Lookup', 'D16:D28').value))
    vid = '|'.join(list(Range('Lookup', 'D29:D44').value))

    dynamic = '|'.join(list(Range('Lookup', 'F2:F3').value))
    other_buy = '|'.join(list(Range('Lookup', 'F4').value))

    platform = np.where(data['Placement'].str.contains(mobile) == True, 'Mobile',
                        np.where(data['Placement'].str.contains(tablet) == True, 'Tablet',
                                 np.where(data['Placement'].str.contains(social) == True, 'Social', 'Desktop')))

    creative = np.where(data['Placement'].str.contains(rm) == True, 'Rich Media',
                        np.where(data['Placement'].str.contains(custom) == True, 'Custom',
                                 np.where(data['Placement'].str.contains(rem) == True, 'Remessaging',
                                          np.where(data['Placement'].str.contains(vid) == True, 'Video', 'Standard'))))

    buy = np.where(data['Placement'].str.contains(dynamic) == True, 'dCPM',
                   np.where(data['Placement'].str.contains(other_buy), 'Flat', ''))

    data['Platform'] = platform
    data['P_Creative'] = creative
    data['Buy'] = buy

    data['Category'] = data['Platform'] + ' - ' + data['P_Creative'] + ' - ' + data['Buy']

    data['Category'] = np.where(data['Category'].str[:3] == ' - ', data['Category'].str[3:], data['Category'])
    data['Category'] = np.where(data['Category'].str[-3:] == ' - ', data['Category'].str[:-3], data['Category'])

    # Message Bucket, Category and Offer
    # 90% of the time, the message bucket, category and offer can be determined from the creative field 1 column. It
    # follows a pattern of Bucket_Category_Offer
    data['Creative Field 1'] = data['Creative Field 1'].str.replace('Creative Type: ', '')
    data['Creative Field 1'] = data['Creative Field 1'].str.replace('(', '')
    data['Creative Field 1'] = data['Creative Field 1'].str.replace(')', '')

    # If Creative Field 1 is equal to (not set), this is either a 1x1 or a placement with logo creative. (not set)
    # fields are therefore set as 'TMO Unique Creative', which is how this has been handled historically.
    data['Creative Field 1'] = data['Creative Field 1'].str.replace('not set', 'TMO Unique Creative')

    # Message Bucket is determined by splitting Creative Field 1 and taking the first word.
    data['Message Bucket'] = data['Creative Field 1'].str.split('_').str.get(0)

    # Message Category is determined by splitting Creative Field 1 and taking the second word.
    data['Message Category'] = data['Creative Field 1'].str.split('_').str.get(1)

    # Message Offer is determined by splitting Creative Field 1 and taking the third word. If the offer is not set,
    # it can sometimes be found in the Creative Groups 2 column. For blanks in the Message Offer column, it will try
    # to pull in the offer from the Creative Groups 2 column.
    data['Message Offer'] = data['Creative Field 1'].str.split('_').str.get(2)
    data['Message Offer'].fillna(data['Creative Groups 2'], inplace=True)

    return data