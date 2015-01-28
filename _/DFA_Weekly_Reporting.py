
# coding: utf-8

# In[69]:

import pandas as pd
import numpy as np
from xlwings import Workbook, Range
import re
import itertools


# In[70]:

wb = Workbook("C:/Users/aarschle1/Google Drive/Optimedia/T-Mobile/Projects/Weekly_Reporting/Opti_DFA_Weekly_Reporting.xlsm")


# In[5]:

sa2 = pd.DataFrame(pd.read_excel(wb.fullname, 'SA_Temp', index_col = None))
cfv2 = pd.DataFrame(pd.read_excel(wb.fullname, 'CFV_Temp', index_col = None))


# In[89]:

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


# In[127]:

sa = sa2
cfv = cfv2


# In[128]:

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


# In[129]:

appended = sa.append(cfv)


# In[130]:

appended['Media Cost'] = np.where(appended['DBM Cost USD'] != 0, appended['DBM Cost USD'], appended['Media Cost'])
appended.drop('DBM Cost USD', 1, inplace = True)


# In[131]:

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
appended['Click-through URL'] = appended['Click-through URL'].apply(lambda x: str(x).split('%')[0])
appended['Click-through URL'] = appended['Click-through URL'].apply(lambda x: str(x).split('_')[0])
appended['Click-through URL'] = appended['Click-through URL'].str.replace('DWTR', '')


# In[132]:

Range('Lookup', 'H1', index = False).value = appended['Click-through URL']


# In[133]:

#appended = appended.groupby(['Campaign', 'Date', 'Site (DCM)', 'Creative', 'Click-through URL', 'Ad', 'Creative Groups 1',
#                             'Creative Groups 2', 'Creative ID', 'Creative Type', 'Creative Field 1', 'Placement',   
#                             'Placement Cost Structure', 'Floodlight Attribution Type', 'Activity', 'OrderNumber (string)',
#                             'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)'], as_index = False).aggregate(np.sum)


# In[134]:

appended['Plans'].fillna(0, inplace = True)
appended['Services'].fillna(0, inplace = True)
appended['Devices'].fillna(0, inplace = True)
appended['Accessories'].fillna(0, inplace = True)
appended['Orders'].fillna(0, inplace = True)


# In[135]:

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


# In[136]:

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


# In[137]:

mobile = '|'.join(list(Range('Lookup', 'B2:B6').value))
tablet = '|'.join(list(Range('Lookup', 'B7:B9').value))
social = '|'.join(list(Range('Lookup', 'B10:B12').value))

rm = '|'.join(list(Range('Lookup', 'D2:D5').value))
custom = '|'.join(list(Range('Lookup', 'D6:D15').value))
rem = '|'.join(list(Range('Lookup', 'D16:D28').value))

dynamic = '|'.join(list(Range('Lookup', 'F2:F3').value))
other_buy = '|'.join(list(Range('Lookup', 'F4').value))


# In[138]:

platform = np.where(appended['Placement'].str.contains(mobile) == True, 'Mobile',
                    np.where(appended['Placement'].str.contains(tablet) == True, 'Tablet',
                             np.where(appended['Placement'].str.contains(social) == True, 'Social', '')))

creative = np.where(appended['Placement'].str.contains(rm) == True, 'Rich Media',
                    np.where(appended['Placement'].str.contains(custom) == True, 'Custom',
                             np.where(appended['Placement'].str.contains(rem) == True, 'Remessaging', 'Standard')))

buy = np.where(appended['Placement'].str.contains(dynamic) == True, 'dCPM',
               np.where(appended['Placement'].str.contains(other_buy), 'Flat', ''))


# In[139]:

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


# In[140]:

view_based = list(set(view_through).intersection(column_names))
click_based = list(set(click_through).intersection(column_names))
SLV_conversions = list(set(store_locator).intersection(column_names))


# In[141]:

appended['Post-Click Activity'] = appended[click_based].sum(axis=1)
appended['Post-Impression Activity'] = appended[view_based].sum(axis=1)
appended['Store Locator Visits'] = appended[SLV_conversions].sum(axis=1)

appended['Awareness Actions'] = appended['A Actions'] + appended['B Actions']
appended['Consideration Actions'] = appended['C Actions'] + appended['D Actions']
appended['Traffic Actions'] = appended['Awareness Actions'] + appended['Consideration Actions']


# In[142]:

appended['Creative Field 1'] = appended['Creative Field 1'].str.replace('Creative Type: ', '')
appended['Creative Field 1'] = appended['Creative Field 1'].str.replace('(', '')
appended['Creative Field 1'] = appended['Creative Field 1'].str.replace(')', '')
appended['Creative Field 1'] = appended['Creative Field 1'].str.replace('not set', 'TMO Unique Creative')

appended['Message Bucket'] = appended['Creative Field 1'].str.split('_').str.get(0)

appended['Message Category'] = appended['Creative Field 1'].str.split('_').str.get(1)

appended['Message Offer'] = appended['Creative Field 1'].str.split('_').str.get(2)
appended['Message Offer'].fillna(appended['Creative Groups 2'], inplace=True)


# In[143]:

appended['Platform'] = platform
appended['P_Creative'] = creative
appended['Buy'] = buy

appended['Category'] = appended['Platform'] + ' - ' + appended['P_Creative'] + ' - ' + appended['Buy']

appended['Category'] = np.where(appended['Category'].str[:3] == ' - ', appended['Category'].str[3:], appended['Category'])
appended['Category'] = np.where(appended['Category'].str[-3:] == ' - ', appended['Category'].str[:-3], appended['Category'])


# In[144]:

appended['Week'] = appended['Date'].min()
appended['Video Completions'] = 0
appended['Video Views'] = 0

appended['F Tag'] = 0
appended['F Actions'] = 0


# In[145]:

sa_columns = list(sa.columns)


# In[146]:

dimensions = ['Week', 'Date', 'Campaign', 'Site (DCM)', 'Click-through URL', 'F Tag', 'Category', 'Message Bucket', 'Message Category', 
              'Message Offer', 'Creative', 'Ad', 'Creative Groups 1', 'Creative Groups 2', 'Creative ID', 'Creative Type', 
              'Creative Field 1', 'Placement', 'Placement Cost Structure', 'OrderNumber (string)', 'Activity', 'Floodlight Attribution Type',
              'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)']

metrics = ['Media Cost', 'Impressions', 'Clicks', 'Orders', 'Plans', 'Add-a-Line', 'Activations', 'Devices', 'Services', 'Accessories',
           'Prepaid Plans', 'Store Locator Visits', 'A Actions', 'B Actions', 'C Actions', 'D Actions', 'E Actions', 'F Actions', 
           'Awareness Actions', 'Consideration Actions', 'Traffic Actions', 'Post-Click Activity', 'Post-Impression Activity',
           'Video Completions', 'Video Views']


# In[147]:

action_tags = sa_columns[sa_columns.index('DBM Cost USD') + 1:]


# In[148]:

new_columns = dimensions + metrics + action_tags


# In[149]:

new_appended = appended[new_columns]


# In[150]:

Range('working', 'A1').horizontal.value = list(new_appended.columns)
chunk_df(new_appended, 'working', 'A2')


# In[151]:

ftags = pd.DataFrame(Range('F_Tags', 'B1').table.value, columns = Range('F_Tags', 'B1').horizontal.value)
ftags.drop(0, inplace = True)
ftags['Tag Name (Concatenated)'] = ftags['Group Name'] + " : " + ftags['Activity Name']
Range('F_Tags', 'G2', index = False).value = ftags['Tag Name (Concatenated)']


# In[152]:

f_tag_range = Range('working', 'F2').vertical

for cell in f_tag_range:
    url = cell.offset(0, -1).get_address(False, False, False)
    cell.formula = '=IFERROR(INDEX(F_Tags!G:G,MATCH(working!' + url + ',F_Tags!E:E,0)),"na")'
    
new_appended['F Tag'] = Range('working', 'F2').vertical.value


# In[153]:

f_tag_list = []
for i in new_appended['F Tag']:
    for j in new_appended.columns:
        tag = re.search(i, j)
        if tag:
            f_tag_list.append(j)


# In[154]:

f_tag_list = list(set(f_tag_list).intersection(new_appended.columns))
f_conversions = list(set(f_tag_list).intersection(new_appended.columns))


# In[155]:

new_appended['F Actions'] = new_appended[f_conversions].sum(axis=1)


# In[156]:

data_columns = dimensions + metrics + action_tags


# In[157]:

data = new_appended[data_columns]
data.fillna(0, inplace = True)

ftag = data['F Tag'].apply(lambda x: str(x).split(':')[0])


# In[ ]:

Range('Lookup', 'G1').value = pd.Series(ftag)


# In[48]:

Range('data', 'A1').value = list(data.columns)
chunk_df(data, 'data', 'A1')


## New Data + Past Data Merge

# In[49]:

past_data = pd.DataFrame(pd.read_excel(wb.fullname, 'data', index_col = None))


# In[ ]:

appended_data = past_data.append(data)
appended_data.drop_duplicates(inplace = True)


# In[ ]:

appended_data['Impressions'].sum()


# In[ ]:

past_data['Impressions'].sum() + data['Impressions'].sum()


# In[45]:

chunk_df(appended_data, 'Sheet1', 'A1')


# In[46]:

Range('Sheet1', 'A1').value = list(appended_data.columns)


# In[ ]:



