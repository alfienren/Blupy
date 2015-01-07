
# coding: utf-8

##### Load necessary packages

# In[120]:

import pandas as pd
import numpy as np
from xlwings import Workbook, Range, Sheet
import re


##### Open the working Excel sheet

# In[121]:

wb = Workbook("C:/Users/aarschle1/Google Drive/Optimedia/T-Mobile/Projects/Weekly_Reporting/Opti_DFA_Weekly_Reporting.xlsm")


##### A VBA subroutinue will create and add the required data to the sheets "SA_Temp" and "CFV_Temp". Create pandas DataFrames from this data.

# In[122]:

sa = Range("SA_Temp", "A2").table.value
cfv = Range("CFV_Temp", "A2").table.value


##### Set the column names of the DataFrame

# In[123]:

sa = pd.DataFrame(sa, columns = Range("SA_Temp", "A1").horizontal.value)
cfv = pd.DataFrame(cfv, columns = Range("CFV_Temp", "A1").horizontal.value)


##### Append the CFV data to the SA data and fill N/A values with 0.

# In[124]:

appended = sa.append(cfv)
appended = appended.fillna(0)


##### With the appended DataFrame, group the data, i.e. compress it, by each column below.

# In[125]:

appended = appended.groupby(['Campaign', 'Date', 'Site (DFA)', 'Creative', 'Click-through URL', 'Ad', 'Creative Groups 1',
                             'Creative Groups 2', 'Creative ID', 'Creative Type', 'Creative Field 1', 'Placement',   
                             'Placement Cost Structure', 'Device (string)', 'Floodlight Attribution Type', 'OrderNumber (string)',
                             'Plan (string)', 'Service (string)']).aggregate(np.sum)


##### Again, fill in N/A values with 0

# In[126]:

appended = appended.fillna(0)


##### Copy the new DataFrame into the Excel sheet on the working tab. Create a new DataFrame off of this data

# In[127]:

Range('working', 'A1').value = appended
appended = pd.DataFrame(Range('working', 'A2').table.value, columns=Range('working', 'A1').horizontal.value)
Range('working', 'A1').value = appended


##### Using the list of actions in the 'Action Reference' tab of the Excel sheet, set lists for each action category.

# In[128]:

a_actions = Range('Action_Reference', 'A2').vertical.value
b_actions = Range('Action_Reference', 'B2').vertical.value
c_actions = Range('Action_Reference', 'C2').vertical.value
d_actions = Range('Action_Reference', 'D2').vertical.value
e_actions = Range('Action_Reference', 'E2').vertical.value

col_head = Range('working', 'A1').horizontal.value


##### Set the actions to lists and search the DataFrame columns for each one, summing each value when found.

# In[129]:

a_actions = list(set(a_actions).intersection(col_head))
b_actions = list(set(b_actions).intersection(col_head))
c_actions = list(set(b_actions).intersection(col_head))
d_actions = list(set(d_actions).intersection(col_head))
e_actions = list(set(e_actions).intersection(col_head))


# In[130]:

view_through = []
i = iter(view_through)
for item in col_head:
    view = re.search('View-through Conversions', item)
    if view:
        view_through.append(item)
        i.next()

click_through = []
j = iter(click_through)
for item in col_head:
    click = re.search('Click-through Conversions', item)
    if click:
        click_through.append(item)
        j.next()


# In[131]:

view_based = list(set(view_through).intersection(col_head))
click_based = list(set(click_through).intersection(col_head))


##### Add columns to the DataFrame for each action category

# In[132]:

appended['A Actions'] = appended[a_actions].sum(axis=1)
appended['B Actions'] = appended[b_actions].sum(axis=1)
appended['C Actions'] = appended[c_actions].sum(axis=1)
appended['D Actions'] = appended[d_actions].sum(axis=1)
appended['E Actions'] = appended[e_actions].sum(axis=1)

appended['Click Based'] = appended[click_based].sum(axis=1)
appended['View Based'] = appended[view_based].sum(axis=1)


##### Store Locator

# In[133]:

store_locator = []
k = iter(store_locator)
for item in col_head:
    locator = re.search('Store Locator', item)
    if locator:
        store_locator.append(item)
        k.next()


# In[134]:

SLV_conversions = list(set(store_locator).intersection(col_head))
appended['Store Locator Visits'] = appended[store_locator].sum(axis=1)


##### Traffic Action Totals

# In[135]:

appended['Awareness Actions'] = appended['A Actions'] + appended['B Actions']
appended['Consideration Actions'] = appended['C Actions'] + appended['D Actions']
appended['Traffic Actions'] = appended['Awareness Actions'] + appended['Consideration Actions']


##### Message Categories

# In[155]:

appended['Creative Field 1'] = appended['Creative Field 1'].str.replace('Creative Type: ', '')

appended['Creative Field 1'] = appended['Creative Field 1'].str.replace('(', '')
appended['Creative Field 1'] = appended['Creative Field 1'].str.replace(')', '')
appended['Creative Field 1'] = appended['Creative Field 1'].str.replace('not set', 'TMO Unique Creative')

appended['Message Bucket'] = appended['Creative Field 1'].str.split('_').str.get(0)

appended['Message Category'] = appended['Creative Field 1'].str.split('_').str.get(1)

appended['Message Offer'] = appended['Creative Field 1'].str.split('_').str.get(2)
appended['Message Offer'].fillna(appended['Creative Groups 2'], inplace=True)


##### Strip the embedded URL encoding used by BlueKai to get the actual URL.

# In[102]:

appended['Click-through URL'] = appended['Click-through URL'].str.replace('http://analytics.bluekai.com/site/', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('15991\?phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('http://15991\?phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('event%3Dclick&phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('event%3Dclick&phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('aid%3D%eadv!&phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('pid%3D%epid!&phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('cid%3D%ebuy!&phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('crid%3D%ecid!&done', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('pid%3D%25epid!&phint', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3Fcmpid%3DWTR_DD_DDRFCBK_RQLMKXRCUQZ1042%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3F%26csdids%3DADV_DS_%epid!_%eaid!_%ecid!_%eadv!', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('=', '')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%2F', '/')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%3A', ':')
appended['Click-through URL'] = appended['Click-through URL'].str.replace('%23', '#')
appended['Click-through URL'] = appended['Click-through URL'].apply(lambda x: x.split('.html')[0])
appended['Click-through URL'] = appended['Click-through URL'].apply(lambda x: x.split('?')[0])


# In[ ]:



