import numpy as np
import pandas as pd
from xlwings import Range

def categories(data):

    # Categories are broken down by Platform (mobile/tablet, social, desktop), followed by placement creative (Rich
    # Media, Custom, Remessaging, banners, etc.). Lastly, the placement buy type (dCPM, Flat, CPM, etc.)

    # Words to match on are included in the Lookup tab of the Excel sheet.
    # Example output of categories:
    #   Desktop - Standard - dCPM
    #   Mobile - Custom - Flat
    mobile = '|'.join(list(Range('Lookup', 'B2:B12').value))
    tablet = '|'.join(list(Range('Lookup', 'B13:B15').value))
    social = '|'.join(list(Range('Lookup', 'B16:B18').value))

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

    data['Platform'] = np.where((data['Platform'].str.contains(mobile) == True) | (data['Platform'].str.contains(tablet) == True), 'Mobile',
                                  np.where(data['Platform'].str.contains(social) == True), 'Social', 'Desktop')

    data['P_Creative'] = np.where(data['P_Creative'].str.contains(vid) == True, 'Video',
                                  np.where(data['P_Creative'].str.contains(social) == True, np.NaN, 'Display'))

    data['Category Adjusted'] = data['Platform'] + ' - ' + data['P_Creative']

    return data

def sites(data):

    sites = pd.DataFrame(Range('Lookup', 'N1').table.value, columns = Range('Lookup', 'N1').horizontal.value)
    sites.drop(0, inplace = True)

    data = pd.merge(data, sites, left_on= 'Site (DCM)', right_on= 'DFA', how= 'left')
    data.drop('DFA', axis = 1, inplace = True)

    return data