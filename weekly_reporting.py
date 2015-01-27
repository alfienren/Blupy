import pandas as pd
import numpy as np
from xlwings import Workbook, Range
import re

wb = Workbook.caller()

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

sa = pd.DataFrame(pd.read_excel(wb.fullname, 'SA_Temp', index_col = None))
cfv = pd.DataFrame(pd.read_excel(wb.fullname, 'CFV_Temp', index_col = None))

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

a_actions = Range('Action_Reference', 'A2').vertical.value
b_actions = Range('Action_Reference', 'B2').vertical.value
c_actions = Range('Action_Reference', 'C2').vertical.value
d_actions = Range('Action_Reference', 'D2').vertical.value
e_actions = Range('Action_Reference', 'E2').vertical.value

traffic_tags = a_actions + b_actions + c_actions + d_actions + e_actions

data = sa.append(cfv)

column_names = data.columns

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

data = data.groupby(['Campaign', 'Date', 'Site (DCM)', 'Creative', 'Click-through URL', 'Ad', 'Creative Groups 1',
                             'Creative Groups 2', 'Creative ID', 'Creative Type', 'Creative Field 1', 'Placement',   
                             'Placement Cost Structure', 'Floodlight Attribution Type', 'Activity', 'OrderNumber (string)',
                             'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)'], as_index = False).aggregate(np.sum)

data['Plans'].fillna(0, inplace = True)
data['Services'].fillna(0, inplace = True)
data['Devices'].fillna(0, inplace = True)
data['Accessories'].fillna(0, inplace = True)

data['Plan (string)'] = np.where(data['Plans'] < 1, '', data['Plan (string)'])

data['Service (string)'] = np.where(data['Services'] < 1, '', data['Service (string)'])

data['Accessory (string)'] = np.where(data['Accessories'] < 1, '', data['Accessory (string)'])

data['Devices'] = np.where(data['Device (string)'].str.contains('nan') == True, 0, data['Devices'])

data['OrderNumber (string)'] = np.where(data['Orders'] < 1, '', data['OrderNumber (string)'])

data['Activity'] = np.where(data['Orders'] < 1, '', data['Activity'])

data['Floodlight Attribution Type'] = np.where(data['Orders'] < 1, '', data['Floodlight Attribution Type'])

data['Devices'] = np.where(data['Device (string)'].str.contains('nan') == True, 0, data['Devices'])

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
data['Store Locator Visits'] = data[store_locator].sum(axis=1)

data['Awareness Actions'] = data['A Actions'] + data['B Actions']
data['Consideration Actions'] = data['C Actions'] + data['D Actions']
data['Traffic Actions'] = data['Awareness Actions'] + data['Consideration Actions']

data['Creative Field 1'] = data['Creative Field 1'].str.replace('Creative Type: ', '')
data['Creative Field 1'] = data['Creative Field 1'].str.replace('(', '')
data['Creative Field 1'] = data['Creative Field 1'].str.replace(')', '')
data['Creative Field 1'] = data['Creative Field 1'].str.replace('not set', 'TMO Unique Creative')

data['Message Bucket'] = data['Creative Field 1'].str.split('_').str.get(0)

data['Message Category'] = data['Creative Field 1'].str.split('_').str.get(1)

data['Message Offer'] = data['Creative Field 1'].str.split('_').str.get(2)
data['Message Offer'].fillna(data['Creative Groups 2'], inplace=True)

data['Week'] = data['Date'].min()
data['Video Completions'] = 0
data['Video Views'] = 0

data['F Tag'] = 0
data['F Actions'] = 0

dimensions = ['Week', 'Date', 'Campaign', 'Site (DCM)', 'Click-through URL', 'F Tag', 'Message Bucket', 'Message Category', 
              'Message Offer', 'Creative', 'Ad', 'Creative Groups 1', 'Creative Groups 2', 'Creative ID', 'Creative Type', 
              'Creative Field 1', 'Placement', 'Placement Cost Structure', 'OrderNumber (string)', 'Activity', 'Floodlight Attribution Type',
              'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)']

metrics = ['Media Cost', 'Impressions', 'Clicks', 'Orders', 'Plans', 'Add-a-Line', 'Activations', 'Devices', 'Services', 'Accessories',
           'Prepaid Plans', 'Store Locator Visits', 'A Actions', 'B Actions', 'C Actions', 'D Actions', 'E Actions', 'F Actions', 
           'Awareness Actions', 'Consideration Actions', 'Traffic Actions', 'Post-Click Activity', 'Post-Impression Activity',
           'Video Completion', 'Video Views']

sa_columns = list(sa.columns)
action_tags = sa_columns[sa_columns.index('DBM Cost USD') + 1:]

columns = dimensions + metrics + action_tags

data = data[columns]

Range('working', 'A1').horizontal.value = list(data.columns)
chunk_df(data, 'working', 'A2')

f_tag_range = Range('working', 'F2').vertical

for cell in f_tag_range:
    url = cell.offset(0, -1).get_address(False, False, False)
    cell.formula = '=IFERROR(INDEX(F_Tags!C:C,MATCH(working!' + url + ',F_Tags!E:E,0)),"na")'

data['F Tag'] = Range('working', 'F2').vertical.value

f_tag_list = []
for i in data['F Tag']:
    for j in data.columns:
        tag = re.search(i, j)
        if tag:
            f_tag_list.append(j)
            
f_tag_list = list(set(f_tag_list).intersection(data.columns))
f_conversions = list(set(f_tag_list).intersection(data.columns))

data['F Actions'] = data[f_conversions].sum(axis=1)