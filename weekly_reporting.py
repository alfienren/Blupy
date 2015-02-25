from __future__ import division

"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
DFA Weekly Reporting
Created by: Aaron Schlegel
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
Load necessary packages
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

from xlwings import Workbook, Range, Sheet
import pandas as pd
import numpy as np
import re

"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
function to load data into Excel without overloading memory
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

def chunk_df(df, sheet, startcell, chunk_size=1000):

    if len(df) <= (chunk_size + 1):
        Range(sheet, startcell, index=False, header=True).value = df

    else:
        Range(sheet, startcell, index=False).value = list(df.columns)
        c = re.match(r"([a-z]+)([0-9]+)", startcell[0] + str(int(startcell[1]) + 1), re.I)
        row = c.group(1)
        col = int(c.group(2))

        for chunk in (df[rw:rw + chunk_size] for rw in
                      range(0, len(df), chunk_size)):
            Range(sheet, row + str(col), index=False, header=False).value = chunk
            col += chunk_size

"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
Main reporting function
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

def dfa_reporting():

    # Before function is ran, VBA code will create the necessary tabs in order to process correctly. See the
    # documentation in the VBA modules for more information.

    wb = Workbook.caller() # Initiate workbook object
    sheet = Range('Lookup', 'AA1').value # Grab workbook path from Excel sheet

    wb.save() # Workbook needs to be saved in order to load the data into pandas properly

    # Load the Site Activity and Custom Floodlight Variable data into pandas as DataFrames

    sa = pd.DataFrame(pd.read_excel(sheet, 'SA_Temp', index_col=None))

    cfv = pd.DataFrame(Range('CFV_Temp', 'A1').table.value, columns = Range('CFV_Temp', 'A1').horizontal.value)
    cfv.drop(0, inplace=True)
    #cfv = pd.DataFrame(pd.read_excel(sheet, 'CFV_Temp', index_col=None))

    cfv['Orders'] = 1 # Create orders column in cfv data. Each OrderNumber counts as 1 order
    cfv['Plans'] = np.where(cfv['Plan (string)'] != '', cfv['Plan (string)'].str.count(',') + 1, 0) # Count the number of plans in the Plans column
    cfv['Services'] = np.where(cfv['Service (string)'] != '', cfv['Service (string)'].str.count(',') + 1, 0) # Count number of services in the Service column
    cfv['Accessories'] = np.where(cfv['Accessory (string)'] != '', cfv['Accessory (string)'].str.count(',') + 1, 0) # Count number of Accessories in the Accessories column
    cfv['Devices'] = np.where(cfv['Device (string)'] != '', cfv['Device (string)'].str.count(',') + 1, 0) # Count number of devices in the Plans column
    cfv['Add-a-Line'] = cfv['Service (string)'].str.count('ADD') # Count number of Add-a-Lines in the Service column
    cfv['Activations'] = cfv['Plans'] + cfv['Add-a-Line'] #Activations are defined as the sum of Plans and Add-a-Line

    # Postpaid plans are defined as a Plan + Device. By row, if number of plans is equal to number of devices, Postpaid
    # plans = number of plans. If plans and devices are not equal, use the minimum number.
    cfv['Postpaid Plans'] = np.where(cfv['Plans'] == cfv['Devices'], cfv['Plans'],
                                     pd.concat([cfv['Plans'], cfv['Devices']], axis=1).min(axis=1))

    # Prepaid plans are defined as the number of Devices with no service and plan. If number of plans and services are
    # 0, count of devices is prepaid. If service and plan are not equal, prepaid plans = 0.
    cfv['Prepaid Plans'] = np.where((cfv['Plans'] == 0) & (cfv['Services'] == 0), cfv['Devices'],
                                    np.where(cfv['Devices'] > (cfv['Plans'] & cfv['Services']),
                                             cfv['Devices'] - pd.concat([cfv['Plans'], cfv['Services']], axis=1).max(axis=1), 0))

    # The DDR campaign counts view-through order credit at 50%. If the campaign name contains 'DDR' and the Floodlight
    # Attribution Type is View-through, the order is multiplied by 0.5.
    cfv['Orders'] = np.where(((cfv['Campaign'].str.contains('DDR') == True) | (cfv['Campaign'].str.contains('Q1_Brand Remessaging') == True)) &
                             (cfv['Floodlight Attribution Type'].str.contains('View-through') == True),
                             cfv['Orders'] * 0.5, cfv['Orders'])

    # Estimated Gross Adds are calculated as the count of Devices with 50% view-through credit.
    # If Floodlight Attribution Type is equal to View-through, the count of Devices is multiplied by 0.5
    cfv['eGAs'] = np.where(cfv['Floodlight Attribution Type'].str.contains('View-through') == True,
                            (cfv['Device (string)'].str.count(',') + 1) / 2,
                            cfv['Device (string)'].str.count(',') + 1)

    """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DDR specific reporting
    """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

    devices = cfv['Device (string)'].str.split(',').apply(pd.Series).stack()
    devices.index = devices.index.droplevel(-1)
    devices.name = "Device IDs"

    cfv_new = cfv[cfv.columns[0:17]].join(devices)
    cfv_new = cfv.append(cfv_new)

    ddr = pd.DataFrame(pd.read_excel(sheet[:sheet.rindex('\\')] + '\\_\\devices.xlsx', 'Device Lookup'))
    excluded_devices = str(Range('Lookup', 'L2').value)
    cfv_new = pd.merge(cfv_new, ddr, how = 'left', left_on = 'Device IDs', right_on = 'Key')

    cfv_new['Prepaid GAs'] = np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                        (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Subcategory'].str.contains('Prepaid') == True) &
                                        (cfv_new['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5, np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) & (cfv_new['Device IDs'].notnull() == True) &
                                                 (cfv_new['Product Subcategory'].str.contains('Prepaid') == True) &
                                                  (cfv_new['Floodlight Attribution Type'].str.contains('Click-through') == True)), 1, 0))

    cfv_new['Postpaid GAs'] = np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                          (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Subcategory'].str.contains('Postpaid') == True) &
                                            (cfv_new['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5, np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                                      (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Subcategory'].str.contains('Postpaid') == True)), 1, 0))

    cfv_new['DDR GAs'] = cfv_new['Postpaid GAs'] + cfv_new['Prepaid GAs']

    cfv_new['Prepaid SIMs'] = np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                         (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Category'].str.contains('SIM card') == True) &
                                            (cfv_new['Product Subcategory'].str.contains('Prepaid') == True) &
                                            (cfv_new['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5, np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                                       (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Category'].str.contains('SIM card') == True) &
                                                        (cfv_new['Product Subcategory'].str.contains('Prepaid') == True) &
                                                       (cfv_new['Floodlight Attribution Type'].str.contains('Click-through') == True)), 1, 0))

    cfv_new['Postpaid SIMs'] = np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                          (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Category'].str.contains('SIM card') == True) &
                                            (cfv_new['Floodlight Attribution Type'].str.contains('View-through') == True) &
                                            (cfv_new['Product Subcategory'].str.contains('Postpaid') == True)), 0.5, np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                                       (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Category'].str.contains('SIM card') == True) &
                                                        (cfv_new['Product Subcategory'].str.contains('Postpaid') == True) &
                                                       (cfv_new['Floodlight Attribution Type'].str.contains('Click-through') == True)), 1, 0))

    cfv_new['Prepaid Mobile Internet'] = np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                                    (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Category'].str.contains('Mobile Internet') == True) &
                                                    (cfv_new['Product Subcategory'].str.contains('Prepaid') == True) &
                                                    (cfv_new['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5, np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                                               (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Category'].str.contains('Mobile Internet') == True) &
                                                                (cfv_new['Product Subcategory'].str.contains('Prepaid') == True)), 1, 0))

    cfv_new['Postpaid Mobile Internet'] = np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                                     (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Category'].str.contains('Mobile Internet') == True) &
                                                        (cfv_new['Product Subcategory'].str.contains('Postpaid') == True) &
                                                        (cfv_new['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5, np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                                                   (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Category'].str.contains('Mobile Internet') == True) &
                                                                    (cfv_new['Product Subcategory'].str.contains('Postpaid') == True) &
                                                                   (cfv_new['Floodlight Attribution Type'].str.contains('Click-through') == True)), 1, 0))

    cfv_new['Prepaid Phone'] = np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                          (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Category'].str.contains('Smartphone') == True) &
                                            (cfv_new['Product Subcategory'].str.contains('Prepaid') == True) &
                                            (cfv_new['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5, np.where((((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                                       (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Category'].str.contains('Smartphone') == True) &
                                                        (cfv_new['Product Subcategory'].str.contains('Prepaid') == True) &
                                                       (cfv_new['Floodlight Attribution Type'].str.contains('Click-through') == True))), 1, 0))

    cfv_new['Postpaid Phone'] = np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) &
                                           (cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Category'].str.contains('Smartphone') == True) &
                                            (cfv_new['Product Subcategory'].str.contains('Postpaid') == True) &
                                            (cfv_new['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5, np.where(((cfv_new['Device IDs'].notnull() == True) & (cfv_new['Product Category'].str.contains('Smartphone') == True) &
                                             (cfv_new['Product Subcategory'].str.contains('Postpaid') == True) &
                                                      (cfv_new['Floodlight Attribution Type'].str.contains('Click-through') == True)), 1, 0))

    cfv_new['DDR New Devices'] = np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) & (cfv_new['Device IDs'].notnull() == True) &
                                 (cfv_new['Activity'].str.contains('New TMO Order') == True) &
                                 (cfv_new['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5, np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) & (cfv_new['Device IDs'].notnull() == True) &
                                          (cfv_new['Activity'].str.contains('New TMO Order') == True) &
                                           (cfv_new['Floodlight Attribution Type'].str.contains('Click-through') == True)), 1, 0))

    cfv_new['DDR Add-a-Line'] = np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) & (cfv_new['Device IDs'].notnull() == True) &
                                    (cfv_new['Activity'].str.contains('New My.TMO Order') == True) &
                                    (cfv_new['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5,
                                     np.where(((cfv_new['Device IDs'].str.contains(excluded_devices) == False) & (cfv_new['Device IDs'].notnull() == True) &
                                              (cfv_new['Activity'].str.contains('New My.TMO Order') == True) &
                                               (cfv_new['Floodlight Attribution Type'].str.contains('Click-through') == True)), 1, 0))

    Range('Sheet16', 'A1').value = cfv_new

    # Append the Custom Floodlight Variable data to the Site Activity data. Columns with matching names are merged
    # together. Columns without matching names are added to the new dataframe.
    data = sa.append(cfv_new)

    campaign_specific_lookup = pd.DataFrame(Range('Lookup', 'H3').table.value, columns= Range('Lookup', 'H3').horizontal.value)
    campaign_specific_lookup.drop(0, inplace=True)
    campaign_specific_lookup.drop('Campaign', axis = 1, inplace = True)

    data = pd.merge(data, campaign_specific_lookup, how = 'left', left_on='Placement', right_on='Placement_category')

    # The DFA field DBM Cost is more accurate for placements using dynamic bidding. If a placement is not using
    # dynamic bidding, DBM Cost = 0. Therefore, if DBM cost does not equal 0, replace the row's media cost with
    # DBM cost. If DBM Cost = 0, Media Cost stays the same.
    data['Media Cost'] = np.where(data['DBM Cost USD'] != 0, data['DBM Cost USD'], data['Media Cost'])

    # Adjust spend to Net to Client
    data['Media Cost'] = data['Media Cost'] / .96759

    # DBM Cost column is then removed as it is no longer needed.
    data.drop('DBM Cost USD', 1, inplace=True)

    # Actual URLs are wrapped in BlueKai calls when outputted from DFA. The following finds and removes the code from
    # the URLs, leaving the cleaned URLs remaining.
    data['Click-through URL'] = data['Click-through URL'].str.replace('http://analytics.bluekai.com/site/', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3F%3DADV_DS_%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('15991\?phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('http://15991\?phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('event%3Dclick&phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('aid%3D%eadv!&phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('pid%3D%epid!&phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('cid%3D%ebuy!&phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('crid%3D%ecid!&done', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('pid%3D%25epid!&phint', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3D%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%26csdids', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('DADV_DS_ADDDVL4Q_EMUL7Y9E1YA4116', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('DWTR_DD_DDRDSPLYPR_JM2694TSP3U5895', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('DWTR_DD_DDRFCBK_RQLMKXRCUQZ1042', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3Fcmpid%3', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('b/refmh_', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3Fcm_mmc%3DPurchase-_-Display-_-Revere-_-Revere', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3Fcm_mmc%3DDisplay-_-Purchase-_-GM-_-Tablet_Base', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3Fcmpid%3DWTR_DD_DDRFCBK_RQLMKXRCUQZ1042%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3F%26csdids%3DADV_DS_%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3Fcm_mmc%3DPurchase-_-Display-_-Revere-_-Revere%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3Fcm_mmc%3DDisplay-_-Purchase-_-GM-_-Tablet_Base%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace(
        '%3Fcmpid%3DWTR_DD_DDRDSPLYPR_JM2694TSP3U5895%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('&csdids%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('=', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%2F', '/')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3A', ':')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%23', '#')
    data['Click-through URL'] = data['Click-through URL'].apply(lambda x: str(x).split('.html')[0])
    data['Click-through URL'] = data['Click-through URL'].apply(lambda x: str(x).split('?')[0])
    data['Click-through URL'] = data['Click-through URL'].apply(lambda x: str(x).split('%')[0])
    data['Click-through URL'] = data['Click-through URL'].apply(lambda x: str(x).split('_')[0])
    data['Click-through URL'] = data['Click-through URL'].str.replace('DWTR', '')

    # CFV columns for Plans, Services, etc. that were created earlier have blank values replaced with 0.
    data['Plans'].fillna(0, inplace=True)
    data['Services'].fillna(0, inplace=True)
    data['Devices'].fillna(0, inplace=True)
    data['Accessories'].fillna(0, inplace=True)
    data['Orders'].fillna(0, inplace=True)

    # The following lines ensure a plan, service, accessory, device, order number, etc. counts are tied to an actual
    # floodlight string.

    # If the count of plans, services, accessories, devices, or orders is less than 1, the string is set to blank. If
    # the count is 1 or greater, the string is associated to the count.
    data['Plan (string)'] = np.where(data['Plans'] < 1, '', data['Plan (string)'])

    data['Service (string)'] = np.where(data['Services'] < 1, '', data['Service (string)'])

    data['Accessory (string)'] = np.where(data['Accessories'] < 1, '', data['Accessory (string)'])

    data['Device (string)'] = np.where(data['Devices'] < 1, "", data['Device (string)'])

    data['OrderNumber (string)'] = np.where(data['Orders'] < 1, '', data['OrderNumber (string)'])

    data['Activity'] = np.where(data['Orders'] < 1, '', data['Activity'])

    data['Floodlight Attribution Type'] = np.where(data['Orders'] < 1, '', data['Floodlight Attribution Type'])

    data['Devices'] = np.where(data['Device (string)'].str.contains('nan') == True, 0, data['Devices'])

    # Take the action tag names from the Action_Reference tab in the Excel sheet for A - E actions.
    a_actions = Range('Action_Reference', 'A2').vertical.value
    b_actions = Range('Action_Reference', 'B2').vertical.value
    c_actions = Range('Action_Reference', 'C2').vertical.value
    d_actions = Range('Action_Reference', 'D2').vertical.value
    e_actions = Range('Action_Reference', 'E2').vertical.value

    # Set the data column names to a variable
    column_names = data.columns

    # For each action tag category (A - E), search the column names to find the action tag. A new list of the A - E
    # actions are compiled from the matches in the column names.
    a_actions = list(set(a_actions).intersection(column_names))
    b_actions = list(set(b_actions).intersection(column_names))
    c_actions = list(set(c_actions).intersection(column_names))
    d_actions = list(set(d_actions).intersection(column_names))
    e_actions = list(set(e_actions).intersection(column_names))

    # With the references set for each action tag category, sum the tags along rows and create columns for each action
    # tag bucket.
    data['A Actions'] = data[a_actions].sum(axis=1)
    data['B Actions'] = data[b_actions].sum(axis=1)
    data['C Actions'] = data[c_actions].sum(axis=1)
    data['D Actions'] = data[d_actions].sum(axis=1)
    data['E Actions'] = data[e_actions].sum(axis=1)

    # Placement Categories
    # Due to lack of formalized naming conventions, the following builds categories for placements based on words
    # contained in the placement name.

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

    # Similar to the logic to get the sum of action tags for each respective category, the following finds the sum
    # of post-click and impression activity as well as Store Locator Visits.

    # For post-view and click activity, the column names of the data are searched for columns containing 'View-through',
    # 'Click-through' or 'Store Locator'. Column matches are then stored in lists.
    view_through = []
    for item in column_names:
        view = re.search('View-through Conversions', item)
        if view:
            view_through.append(item)

    click_through = []
    for item in column_names:
        click = re.search('Click-through Conversions', item)
        if click:
            click_through.append(item)

    store_locator = []
    for item in column_names:
        locator = re.search('Store Locator 2', item)
        if locator:
            store_locator.append(item)

    # The matches against the column names are then set
    view_based = list(set(view_through).intersection(column_names))
    click_based = list(set(click_through).intersection(column_names))
    SLV_conversions = list(set(store_locator).intersection(column_names))

    # With the matching references set, sum the matches for post-click and impression activity as well as SLV by row,
    # creating columns for each.
    data['Post-Click Activity'] = data[click_based].sum(axis=1)
    data['Post-Impression Activity'] = data[view_based].sum(axis=1)
    data['Store Locator Visits'] = data[SLV_conversions].sum(axis=1)

    # Create columns for Awareness and Consideration actions. Awareness Actions are the sum of A and B actions,
    # Consideration is the sum of C and D actions
    # Traffic Actions are the total of Awareness and Consideration Actions, or A - D actions.
    data['Awareness Actions'] = data['A Actions'] + data['B Actions']
    data['Consideration Actions'] = data['C Actions'] + data['D Actions']
    data['Traffic Actions'] = data['Awareness Actions'] + data['Consideration Actions']

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

    # Create week column by taking the oldest date in the data
    data['Week'] = data['Date'].min()

    # Add columns for Video Completions and Views, primarily for compatibility between campaigns that run video and
    # don't run video. Those that don't can keep the columns set to zero, but those that have video can then be
    # adjusted with passback data.
    data['Video Completions'] = 0
    data['Video Views'] = 0

    # Similar to the video columns, add columns for F Tag and F Actions.
    data['F Tag'] = 0
    data['F Actions'] = 0

    sa_columns = list(sa.columns)

    # the dimensions variable is used to tell what data we want to keep. Each string below represents the column
    # that will be outputted.
    dimensions = ['Week', 'Date', 'Campaign', 'Site (DCM)', 'Click-through URL', 'F Tag', 'Category', 'Message Bucket',
                  'Message Category', 'Creative Type', 'Creative Groups 1', 'Creative ID', 'Message Offer', 'Creative',
                  'Ad', 'Creative Groups 2',
                  'Creative Field 1', 'Placement', 'Placement ID', 'Placement Cost Structure', 'OrderNumber (string)', 'Activity',
                  'Floodlight Attribution Type',
                  'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)']


    # Similar to the dimensions variable, metrics lists all the data we want to be outputted.
    metrics = ['Media Cost', 'Impressions', 'Clicks', 'Orders', 'Plans', 'Add-a-Line', 'Activations', 'Devices',
               'Services', 'Accessories',
               'Postpaid Plans', 'Prepaid Plans', 'eGAs', 'Store Locator Visits', 'A Actions', 'B Actions', 'C Actions',
               'D Actions', 'E Actions', 'F Actions',
               'Awareness Actions', 'Consideration Actions', 'Traffic Actions', 'Post-Click Activity',
               'Post-Impression Activity', 'Video Completions', 'Video Views', 'Prepaid GAs', 'Postpaid GAs',
               'Prepaid SIMs', 'Postpaid SIMs', 'Prepaid Mobile Internet', 'Postpaid Mobile Internet', 'Prepaid Phone',
               'Postpaid Phone', 'DDR New Devices', 'DDR Add-a-Line']

    # Get all the action tag names so these can be included in the outputted data as well.
    action_tags = sa_columns[sa_columns.index('DBM Cost USD') + 1:]

    # new_columns is set as the a master list of the dimensions, metrics and action tags. The data will be outputted
    # in this order.
    new_columns = dimensions + metrics + action_tags

    # A new DataFrame is created with the columns, removing any extraneous or working data that was used and making
    # sure the order of the data is consistent.
    data = data[new_columns]

    # Copy the new DataFrame into the working tab of the Excel worksheet.
    chunk_df(data, 'working', 'A1')

    # Create a DataFrame of the F Tag sheet included in the Excel worksheet.
    ftags = pd.DataFrame(Range('F_Tags', 'B1').table.value, columns=Range('F_Tags', 'B1').horizontal.value)
    ftags.drop(0, inplace=True)

    # Add a new column to the DataFrame to concatenate the Group Name of the tag with the Activity Name. This will
    # give us a reference we can use to match to the tag to the data.
    ftags['Tag Name (Concatenated)'] = ftags['Group Name'] + " : " + ftags['Activity Name']
    Range('F_Tags', 'G2', index=False).value = ftags['Tag Name (Concatenated)']

    # The F Tag Range is set as column F in the working data (The F Tag column)
    f_tag_range = Range('working', 'F2').vertical

    # for each cell in the range, an INDEX + MATCH formula is entered to find the F Tag for the URL listed in the column
    # to the left.
    for cell in f_tag_range:
        url = cell.offset(0, -1).get_address(False, False, False)
        cell.formula = '=IFERROR(INDEX(F_Tags!G:G,MATCH(working!' + url + ',F_Tags!E:E,0)),"na")'

    # With the F Tag names entered, update the DataFrame's F Tag column with the inputted data.
    data['F Tag'] = Range('working', 'F2').vertical.value

    # the F Tags names that were inputted are then matched to the headers of the columns. When a match is found, the
    # reference is set in a list.
    f_tags = []
    for i in data['F Tag']:
        for j in data.columns:
            tag = re.search(i, j)
            if tag:
                f_tags.append(j)

    # After all the F Tag names have been iterated through to find the appropriate tag columns, the references are then
    # set as the intersection of the column names.
    f_tags = list(set(f_tags).intersection(data.columns))
    f_conversions = list(set(f_tags).intersection(data.columns))

    # The F Actions column that was created earlier is then updated with the sum of the F Action by row based on the
    # corresponding columns to that tag.
    data['F Actions'] = data[f_conversions].sum(axis=1)

    # Strip everything before the colon in the F Tag column to remove the group name.
    data['F Tag'] = data['F Tag'].apply(lambda x: str(x).split(':')[-1])

    # New variable to set the desired columns for the new updated data.
    data_columns = dimensions + metrics + action_tags

    data = data[data_columns]
    data.fillna(0, inplace=True)

    # Workbook is then saved in order for data shuttling to work correctly.
    wb.save()

    # If the data tab in the Excel workbook is blank, the cleaned data will just be placed into the tab.
    if Range('data', 'A1').value is None:
        chunk_df(data, 'data', 'A1')

    # If data is already present in the tab, the two data sets are merged together and then copied into the data tab.
    else:
        past_data = pd.DataFrame(pd.read_excel(sheet, 'data', index_col=None))
        appended_data = past_data.append(data)
        appended_data = appended_data[data_columns]
        appended_data.fillna(0, inplace=True)
        Sheet('data').clear()
        chunk_df(appended_data, 'data', 'A1')