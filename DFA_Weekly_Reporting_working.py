
# coding: utf-8

# In[1]:

import pandas as pd
import numpy as np
from xlwings import Workbook, Range
import re
import itertools


# In[2]:

wb = Workbook("C:/Users/aarschle1/Google Drive/Optimedia/T-Mobile/Projects/Weekly_Reporting/Opti_DFA_Weekly_Reporting.xlsm")


# In[3]:

sa2 = pd.DataFrame(pd.read_excel(wb.fullname, 'SA_Temp', index_col = None))
cfv2 = pd.DataFrame(pd.read_excel(wb.fullname, 'CFV_Temp', index_col = None))


# In[71]:

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


# In[72]:

sa = sa2
cfv = cfv2


# In[73]:

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


# In[106]:

appended = sa.append(cfv)
merged = pd.merge(sa, cfv, on = ['Campaign', 'Date', 'Site (DCM)', 'Creative', 'Click-through URL', 'Ad', 'Creative Groups 1',
                                 'Creative Groups 2', 'Creative ID', 'Creative Type', 'Placement'], left_index = True, how = 'left')


# In[107]:

merged['Impressions'].sum()


# In[75]:

appended['Media Cost'] = np.where(appended['DBM Cost USD'] != 0, appended['DBM Cost USD'], appended['Media Cost'])
appended.drop('DBM Cost USD', 1, inplace = True)


# In[76]:

appended['Click-through URL'] = appended['Click-through URL'].str.replace('http://analytics.bluekai.com/site/', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3F%3DADV_DS_%epid!_%eaid!_%ecid!_%eadv!', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('15991\?phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('http://15991\?phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('event%3Dclick&phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('aid%3D%eadv!&phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('pid%3D%epid!&phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('cid%3D%ebuy!&phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('crid%3D%ecid!&done', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('pid%3D%25epid!&phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3D%epid!_%eaid!_%ecid!_%eadv!', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%26csdids', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('DADV_DS_ADDDVL4Q_EMUL7Y9E1YA4116', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('DWTR_DD_DDRDSPLYPR_JM2694TSP3U5895', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('DWTR_DD_DDRFCBK_RQLMKXRCUQZ1042', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3Fcmpid%3', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('b/refmh_', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3Fcm_mmc%3DPurchase-_-Display-_-Revere-_-Revere', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3Fcm_mmc%3DDisplay-_-Purchase-_-GM-_-Tablet_Base', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3Fcmpid%3DWTR_DD_DDRFCBK_RQLMKXRCUQZ1042%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3F%26csdids%3DADV_DS_%epid!_%eaid!_%ecid!_%eadv!', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3Fcm_mmc%3DPurchase-_-Display-_-Revere-_-Revere%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3Fcm_mmc%3DDisplay-_-Purchase-_-GM-_-Tablet_Base%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3Fcmpid%3DWTR_DD_DDRDSPLYPR_JM2694TSP3U5895%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('&csdids%epid!_%eaid!_%ecid!_%eadv!', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('=', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%2F', '/')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3A', ':')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%23', '#')
appended['Click-through URL'] = appended['Click-through URL'].apply(lambda x: str(x).split('.html')[0])
appended['Click-through URL'] = appended['Click-through URL'].apply(lambda x: str(x).split('?')[0])
appended['Click-through URL'] = appended['Click-through URL'].apply(lambda x: str(x).split('%')[0])


# In[77]:

appended = appended.groupby(['Campaign', 'Date', 'Site (DCM)', 'Creative', 'Click-through URL', 'Ad', 'Creative Groups 1',
                             'Creative Groups 2', 'Creative ID', 'Creative Type', 'Creative Field 1', 'Placement',   
                             'Placement Cost Structure', 'Floodlight Attribution Type', 'Activity', 'OrderNumber (string)',
                             'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)'], as_index = False).aggregate(np.sum)


# In[78]:

appended['Plans'].fillna(0, inplace = True)
appended['Services'].fillna(0, inplace = True)
appended['Devices'].fillna(0, inplace = True)
appended['Accessories'].fillna(0, inplace = True)
appended['Orders'].fillna(0, inplace = True)


# In[79]:

appended['Plan (string)'] = np.where(appended['Plans'] < 1, '',
                                     appended['Plan (string)'])

appended['Service (string)'] = np.where(appended['Services'] < 1, '',
                                       appended['Service (string)'])

appended['Accessory (string)'] = np.where(appended['Accessories'] < 1, '',
                                          appended['Accessory (string)'])

appended['Device (string)'] = np.where(appended['Devices'] < 1, "",
                                       appended['Device (string)'])

appended['OrderNumber (string)'] = np.where(appended['Orders'] < 1, '', appended['OrderNumber (string)'])

appended['Activity'] = np.where(appended['Orders'] < 1, '', appended['Activity'])

appended['Floodlight Attribution Type'] = np.where(appended['Orders'] < 1, '', appended['Floodlight Attribution Type'])

appended['Devices'] = np.where(appended['Device (string)'].str.contains('nan') == True, 0, appended['Devices'])


# In[80]:

a_actions = Range('Action_Reference', 'A2').vertical.value
b_actions = Range('Action_Reference', 'B2').vertical.value
c_actions = Range('Action_Reference', 'C2').vertical.value
d_actions = Range('Action_Reference', 'D2').vertical.value
e_actions = Range('Action_Reference', 'E2').vertical.value

column_names = appended.columns
traffic_tags = a_actions + b_actions + c_actions + d_actions + e_actions

a_actions = list(set(a_actions).intersection(column_names))
b_actions = list(set(b_actions).intersection(column_names))
c_actions = list(set(c_actions).intersection(column_names))
d_actions = list(set(d_actions).intersection(column_names))
e_actions = list(set(e_actions).intersection(column_names))

appended['A Actions'] = appended[a_actions].sum(axis=1)
appended['B Actions'] = appended[b_actions].sum(axis=1)
appended['C Actions'] = appended[c_actions].sum(axis=1)
appended['D Actions'] = appended[d_actions].sum(axis=1)
appended['E Actions'] = appended[e_actions].sum(axis=1)


# In[81]:

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


# In[82]:

view_based = list(set(view_through).intersection(column_names))
click_based = list(set(click_through).intersection(column_names))
SLV_conversions = list(set(store_locator).intersection(column_names))


# In[83]:

appended['Post-Click Activity'] = appended[click_based].sum(axis=1)
appended['Post-Impression Activity'] = appended[view_based].sum(axis=1)
appended['Store Locator Visits'] = appended[store_locator].sum(axis=1)

appended['Awareness Actions'] = appended['A Actions'] + appended['B Actions']
appended['Consideration Actions'] = appended['C Actions'] + appended['D Actions']
appended['Traffic Actions'] = appended['Awareness Actions'] + appended['Consideration Actions']


# In[84]:

appended['Creative Field 1'] = appended['Creative Field 1'].str.replace('Creative Type: ', '')
appended['Creative Field 1'] = appended['Creative Field 1'].str.replace('(', '')
appended['Creative Field 1'] = appended['Creative Field 1'].str.replace(')', '')
appended['Creative Field 1'] = appended['Creative Field 1'].str.replace('not set', 'TMO Unique Creative')

appended['Message Bucket'] = appended['Creative Field 1'].str.split('_').str.get(0)

appended['Message Category'] = appended['Creative Field 1'].str.split('_').str.get(1)

appended['Message Offer'] = appended['Creative Field 1'].str.split('_').str.get(2)
appended['Message Offer'].fillna(appended['Creative Groups 2'], inplace=True)


# In[85]:

appended['Week'] = appended['Date'].min()
appended['Video Completions'] = 0
appended['Video Views'] = 0

appended['F Tag'] = 0
appended['F Actions'] = 0


# In[86]:

sa_columns = list(sa.columns)


# In[87]:

dimensions = ['Week', 'Date', 'Campaign', 'Site (DCM)', 'Click-through URL', 'F Tag', 'Message Bucket', 'Message Category', 
              'Message Offer', 'Creative', 'Ad', 'Creative Groups 1', 'Creative Groups 2', 'Creative ID', 'Creative Type', 
              'Creative Field 1', 'Placement', 'Placement Cost Structure', 'OrderNumber (string)', 'Activity', 'Floodlight Attribution Type',
              'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)']

metrics = ['Media Cost', 'Impressions', 'Clicks', 'Orders', 'Plans', 'Add-a-Line', 'Activations', 'Devices', 'Services', 'Accessories',
           'Prepaid Plans', 'Store Locator Visits', 'A Actions', 'B Actions', 'C Actions', 'D Actions', 'E Actions', 'F Actions', 
           'Awareness Actions', 'Consideration Actions', 'Traffic Actions', 'Post-Click Activity', 'Post-Impression Activity',
           'Video Completions', 'Video Views']


# In[88]:

action_tags = sa_columns[sa_columns.index('DBM Cost USD') + 1:]


# In[89]:

new_columns = dimensions + metrics + action_tags


# In[90]:

new_appended = appended[new_columns]


# In[91]:

Range('working', 'A1').horizontal.value = list(new_appended.columns)
chunk_df(new_appended, 'working', 'A2')


# In[92]:

f_tag_range = Range('working', 'F2').vertical

for cell in f_tag_range:
    url = cell.offset(0, -1).get_address(False, False, False)
    cell.formula = '=IF(' + url + '="http://www.t-mobile.com/","na",IFERROR(INDEX(F_Tags!C:C,MATCH(working!' + url + ',F_Tags!E:E,0)),"na"))'
    
new_appended['F Tag'] = Range('working', 'F2').vertical.value


# In[93]:

f_tag_list = []
for i in new_appended['F Tag']:
    for j in new_appended.columns:
        tag = re.search(i, j)
        if tag:
            f_tag_list.append(j)


# In[94]:

f_tag_list2 = list(set(f_tag_list).intersection(new_appended.columns))
f_conversions = list(set(f_tag_list2).intersection(new_appended.columns))


# In[95]:

new_appended['F Actions'] = new_appended[f_conversions].sum(axis=1)


# In[139]:

pivot = pd.pivot_table(new_appended, index = dimensions)


# In[140]:

pivot


# In[149]:

Range('data', 'A1').value = list(pivot.index.names)
chunk_df(pivot.index, 'data', 'A2')


# In[150]:

Range('data', 'A1').horizontal.offset(0, 1).value = list(pivot.columns)
#chunk_df(pivot, 'data',)


# In[ ]:



