import pandas as pd
import numpy as np
from xlwings import Workbook, Range, Sheet
import re

def chunk_df(df, sheet, startcell, chunk_size = 10000):
    if len(df) <= (chunk_size + 1):
        Range(sheet, startcell, index = False, header = True).value = df
    else:
        c = re.match(r"([a-z]+)([0-9]+)", startcell, re.I)
        row = c.group(1)
        col = int(c.group(2))
        
        for chunk in (df[rw:rw + chunk_size] for rw in 
                      range(0, len(df), chunk_size)):
            Range(sheet, row + str(col), index = False, header = False).value = chunk
            col += chunk_size
            
def dfa_reporting():
    
    wb = Workbook.caller()
    
    #sa = pd.DataFrame(pd.read_excel(wb, 'SA_Temp', index_col = None))
    #cfv = pd.DataFrame(pd.read_excel(wb, 'CFV_Temp', index_col = None))
    
    sa = pd.DataFrame(Range('SA_Temp', 'A1').table.value, columns = Range('SA_Temp', 'A1').horizontal.value)
    cfv = pd.DataFrame(Range('CFV_Temp', 'A1').table.value, columns = Range('CFV_Temp', 'A1').horizontal.value)
    
    sa.drop(0, inplace = True)    
    cfv.drop(0, inplace = True)
    
    cfv['Orders'] = 1
    cfv['Plans'] = cfv['Plan (string)'].str.count(',') + 1
    cfv['Devices'] = cfv['Device (string)'].str.count(',') + 1
    cfv['Services'] = cfv['Service (string)'].str.count(',') + 1
    cfv['Add-a-Line'] = cfv['Service (string)'].str.count('ADD')
    cfv['Accessories'] = cfv['Accessory (string)'].str.count(',') + 1
    cfv['Activations'] = cfv['Plans'] + cfv['Add-a-Line']
    
    cfv['Postpaid Plans'] = np.where(cfv['Plans'] == cfv['Devices'], cfv['Plans'], pd.concat([cfv['Plans'], cfv['Devices']], axis=1).min(axis=1))
    cfv['Prepaid Plans'] = np.where((cfv['Plans'] == 0) & (cfv['Devices'] != 0), 0, cfv['Devices'])
    
    cfv ['Orders'] = np.where((cfv['Campaign'].str.contains('DDR') == True) & (cfv['Floodlight Attribution Type'].str.contains('View-through') == True),
                               cfv['Orders'] * 0.5, cfv['Orders'])
    
    data = sa.append(cfv)
    
    data['Media Cost'] = np.where(data['DBM Cost USD'] != 0, data['DBM Cost USD'], data['Media Cost'])
    data.drop('DBM Cost USD', 1, inplace = True)

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
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3Fcm_mmc%3DPurchase-_-Display-_-Revere-_-Revere', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3Fcm_mmc%3DDisplay-_-Purchase-_-GM-_-Tablet_Base', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3Fcmpid%3DWTR_DD_DDRFCBK_RQLMKXRCUQZ1042%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3F%26csdids%3DADV_DS_%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3Fcm_mmc%3DPurchase-_-Display-_-Revere-_-Revere%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3Fcm_mmc%3DDisplay-_-Purchase-_-GM-_-Tablet_Base%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
    data['Click-through URL'] = data['Click-through URL'].str.replace('%3Fcmpid%3DWTR_DD_DDRDSPLYPR_JM2694TSP3U5895%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
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

    data = data.groupby(['Campaign', 'Date', 'Site (DCM)', 'Creative', 'Click-through URL', 'Ad', 'Creative Groups 1',
                                 'Creative Groups 2', 'Creative ID', 'Creative Type', 'Creative Field 1', 'Placement',   
                                 'Placement Cost Structure', 'Floodlight Attribution Type', 'Activity', 'OrderNumber (string)',
                                 'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)'], as_index = False).aggregate(np.sum)

    data['Plans'].fillna(0, inplace = True)
    data['Services'].fillna(0, inplace = True)
    data['Devices'].fillna(0, inplace = True)
    data['Accessories'].fillna(0, inplace = True)
    data['Orders'].fillna(0, inplace = True)
    
    data['Plan (string)'] = np.where(data['Plans'] < 1, '', data['Plan (string)'])
    
    data['Service (string)'] = np.where(data['Services'] < 1, '', data['Service (string)'])
    
    data['Accessory (string)'] = np.where(data['Accessories'] < 1, '', data['Accessory (string)'])
    
    data['Device (string)'] = np.where(data['Devices'] < 1, "", data['Device (string)'])
    
    data['OrderNumber (string)'] = np.where(data['Orders'] < 1, '', data['OrderNumber (string)'])
    
    data['Activity'] = np.where(data['Orders'] < 1, '', data['Activity'])
    
    data['Floodlight Attribution Type'] = np.where(data['Orders'] < 1, '', data['Floodlight Attribution Type'])
    
    data['Devices'] = np.where(data['Device (string)'].str.contains('nan') == True, 0, data['Devices'])

    a_actions = Range('Action_Reference', 'A2').vertical.value
    b_actions = Range('Action_Reference', 'B2').vertical.value
    c_actions = Range('Action_Reference', 'C2').vertical.value
    d_actions = Range('Action_Reference', 'D2').vertical.value
    e_actions = Range('Action_Reference', 'E2').vertical.value
    
    column_names = data.columns
    
    a_actions = list(set(a_actions).intersection(column_names))
    b_actions = list(set(b_actions).intersection(column_names))
    c_actions = list(set(c_actions).intersection(column_names))
    d_actions = list(set(d_actions).intersection(column_names))
    e_actions = list(set(e_actions).intersection(column_names))
    
    data['A Actions'] = data[a_actions].sum(axis=1)
    data['B Actions'] = data[b_actions].sum(axis=1)
    data['C Actions'] = data[c_actions].sum(axis=1)
    data['D Actions'] = data[d_actions].sum(axis=1)
    data['E Actions'] = data[e_actions].sum(axis=1)
    
    mobile = '|'.join(list(Range('Lookup', 'B2:B6').value))
    tablet = '|'.join(list(Range('Lookup', 'B7:B9').value))
    social = '|'.join(list(Range('Lookup', 'B10:B12').value))
    
    rm = '|'.join(list(Range('Lookup', 'D2:D5').value))
    custom = '|'.join(list(Range('Lookup', 'D6:D15').value))
    rem = '|'.join(list(Range('Lookup', 'D16:D28').value))
    
    dynamic = '|'.join(list(Range('Lookup', 'F2:F3').value))
    other_buy = '|'.join(list(Range('Lookup', 'F4').value))
    
    platform = np.where(data['Placement'].str.contains(mobile) == True, 'Mobile',
                        np.where(data['Placement'].str.contains(tablet) == True, 'Tablet',
                                 np.where(data['Placement'].str.contains(social) == True, 'Social', '')))
    
    creative = np.where(data['Placement'].str.contains(rm) == True, 'Rich Media',
                        np.where(data['Placement'].str.contains(custom) == True, 'Custom',
                                 np.where(data['Placement'].str.contains(rem) == True, 'Remessaging', 'Standard')))
    
    buy = np.where(data['Placement'].str.contains(dynamic) == True, 'dCPM',
                   np.where(data['Placement'].str.contains(other_buy), 'Flat', ''))
    
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
        locator = re.search('Store Locator', item)
        if locator:
            store_locator.append(item)
    
    view_based = list(set(view_through).intersection(column_names))
    click_based = list(set(click_through).intersection(column_names))
    SLV_conversions = list(set(store_locator).intersection(column_names))
    
    data['Post-Click Activity'] = data[click_based].sum(axis=1)
    data['Post-Impression Activity'] = data[view_based].sum(axis=1)
    data['Store Locator Visits'] = data[SLV_conversions].sum(axis=1)
    
    data['Awareness Actions'] = data['A Actions'] + data['B Actions']
    data['Consideration Actions'] = data['C Actions'] + data['D Actions']
    data['Traffic Actions'] = data['Awareness Actions'] + data['Consideration Actions']

    data['eGAs'] = np.where(data['Floodlight Attribution Type'].str.contains('View-through') == True,
                            (data['Device (string)'].str.count(',') + 1) / 2, 
                                data['Device (string)'].str.count(',') + 1)
                                
    data['Creative Field 1'] = data['Creative Field 1'].str.replace('Creative Type: ', '')
    data['Creative Field 1'] = data['Creative Field 1'].str.replace('(', '')
    data['Creative Field 1'] = data['Creative Field 1'].str.replace(')', '')
    data['Creative Field 1'] = data['Creative Field 1'].str.replace('not set', 'TMO Unique Creative')
    
    data['Message Bucket'] = data['Creative Field 1'].str.split('_').str.get(0)
    
    data['Message Category'] = data['Creative Field 1'].str.split('_').str.get(1)
    
    data['Message Offer'] = data['Creative Field 1'].str.split('_').str.get(2)
    data['Message Offer'].fillna(data['Creative Groups 2'], inplace=True)

    data['Platform'] = platform
    data['P_Creative'] = creative
    data['Buy'] = buy
    
    data['Category'] = data['Platform'] + ' - ' + data['P_Creative'] + ' - ' + data['Buy']
    
    data['Category'] = np.where(data['Category'].str[:3] == ' - ', data['Category'].str[3:], data['Category'])
    data['Category'] = np.where(data['Category'].str[-3:] == ' - ', data['Category'].str[:-3], data['Category'])
    
    data['Week'] = data['Date'].min()
    data['Video Completions'] = 0
    data['Video Views'] = 0
    
    data['F Tag'] = 0
    data['F Actions'] = 0
    
    sa_columns = list(sa.columns)
    
    dimensions = ['Week', 'Date', 'Campaign', 'Site (DCM)', 'Click-through URL', 'F Tag', 'Category', 'Message Bucket', 'Message Category', 
                  'Message Offer', 'Creative', 'Ad', 'Creative Groups 1', 'Creative Groups 2', 'Creative ID', 'Creative Type', 
                  'Creative Field 1', 'Placement', 'Placement Cost Structure', 'OrderNumber (string)', 'Activity', 'Floodlight Attribution Type',
                  'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)']
    
    metrics = ['Media Cost', 'Impressions', 'Clicks', 'Orders', 'Plans', 'Add-a-Line', 'Activations', 'Devices', 'Services', 'Accessories',
               'Prepaid Plans', 'eGAs', 'Store Locator Visits', 'A Actions', 'B Actions', 'C Actions', 'D Actions', 'E Actions', 'F Actions', 
               'Awareness Actions', 'Consideration Actions', 'Traffic Actions', 'Post-Click Activity', 'Post-Impression Activity',
               'Video Completions', 'Video Views']

    action_tags = sa_columns[sa_columns.index('DBM Cost USD') + 1:]
    
    new_columns = dimensions + metrics + action_tags
    
    data = data[new_columns]
    
    Range('working', 'A1').horizontal.value = list(data.columns)
    chunk_df(data, 'working', 'A2')
    
    ftags = pd.DataFrame(Range('F_Tags', 'B1').table.value, columns = Range('F_Tags', 'B1').horizontal.value)
    ftags.drop(0, inplace = True)
    ftags['Tag Name (Concatenated)'] = ftags['Group Name'] + " : " + ftags['Activity Name']
    Range('F_Tags', 'G2', index = False).value = ftags['Tag Name (Concatenated)']

    f_tag_range = Range('working', 'F2').vertical
    
    for cell in f_tag_range:
        url = cell.offset(0, -1).get_address(False, False, False)
        cell.formula = '=IF(' + url + '="http://www.t-mobile.com/","na",IFERROR(INDEX(F_Tags!C:C,MATCH(working!' + url + ',F_Tags!E:E,0)),"na"))'
        
    data['F Tag'] = Range('working', 'F2').vertical.value
    
    f_tags = []
    for i in data['F Tag']:
        for j in data.columns:
            tag = re.search(i, j)
            if tag:
                f_tags.append(j)
    
    f_tags = list(set(f_tags).intersection(data.columns))
    f_conversions = list(set(f_tags).intersection(data.columns))
    
    data['F Actions'] = data[f_conversions].sum(axis=1)
    
    data['F Tag'] = data['F Tag'].apply(lambda x: str(x).split(':')[0])
    
    data_columns = dimensions + metrics + action_tags

    data = data[data_columns]
    data.fillna(0, inplace = True)
    
    wb.save()
    
    past_data = pd.DataFrame(pd.read_excel(wb.fullname, 'data', index_col = None))
    appended_data = past_data.append(data)
    appended_data.drop_duplicates(inplace = True)
    
    Sheet('data').clearcontents()
    
    Range('data', 'A1').value = appended_data.columns
    chunk_df(appended_data, 'data', 'A2')