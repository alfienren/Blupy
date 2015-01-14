
# coding: utf-8

##### Load necessary packages

# In[1]:

import pandas as pd
import numpy as np
from xlwings import Workbook, Range, Sheet
import re
import itertools
from splinter import Browser
from bs4 import BeautifulSoup


##### Open the working Excel sheet

# In[2]:

wb = Workbook("C:/Users/aarschle1/Google Drive/Optimedia/T-Mobile/Projects/Weekly_Reporting/Opti_DFA_Weekly_Reporting.xlsm")


##### A VBA subroutinue will create and add the required data to the sheets "SA_Temp" and "CFV_Temp". Create pandas DataFrames from this data.

# In[3]:

sa = Range("SA_Temp", "A1").table.value
cfv = Range("CFV_Temp", "A1").table.value


##### Set the column names of the DataFrame

# In[4]:

sa = pd.DataFrame(Range("SA_Temp", "A1").table.value, columns = Range("SA_Temp", "A1").horizontal.value)
cfv = pd.DataFrame(Range("CFV_Temp", "A1").table.value, columns = Range("CFV_Temp", "A1").horizontal.value)

sa = sa.drop(0)
cfv = cfv.drop(0)


##### Transform CFV Data

# In[5]:

cfv['Orders'] = 1


# In[6]:

cfv['Plans'] = cfv['Plan (string)'].str.count(',') + 1
cfv['Devices'] = cfv['Device (string)'].str.count(',') + 1
cfv['Services'] = cfv['Service (string)'].str.count(',') + 1
cfv['Add-a-Line'] = cfv['Service (string)'].str.count('ADD')
cfv['Accessories'] = cfv['Accessory (string)'].str.count(',') + 1
cfv['Activations'] = cfv['Plans'] + cfv['Add-a-Line']

cfv['Plans'] = cfv['Plans'].fillna(0)
cfv['Devices'] = cfv['Devices'].fillna(0)
cfv['Services'] = cfv['Services'].fillna(0)
cfv['Add-a-Line'] = cfv['Add-a-Line'].fillna(0)
cfv['Accessories'] = cfv['Accessories'].fillna(0)

postpaid = np.where(cfv['Plans'] == cfv['Devices'], cfv['Plans'], pd.concat([cfv['Plans'], cfv['Devices']], axis=1).min(axis=1))
prepaid = np.where((cfv['Plans'] == 0) & (cfv['Devices'] != 0), 0, cfv['Devices'])

cfv['Postpaid Plans'] = postpaid
cfv['Prepaid Plans'] = prepaid


##### Append the CFV data to the SA data and fill N/A values with 0.

# In[7]:

appended = sa.append(cfv)
appended = appended.fillna(0)


##### With the appended DataFrame, group the data, i.e. compress it, by each column below.

# In[8]:

appended = appended.groupby(['Campaign', 'Date', 'Site (DFA)', 'Creative', 'Click-through URL', 'Ad', 'Creative Groups 1',
                             'Creative Groups 2', 'Creative ID', 'Creative Type', 'Creative Field 1', 'Placement',   
                             'Placement Cost Structure', 'Device (string)', 'Floodlight Attribution Type', 'OrderNumber (string)',
                             'Plan (string)', 'Service (string)']).aggregate(np.sum)


# In[9]:

Range('working', 'A1').value = appended
appended = pd.DataFrame(Range('working', 'A2').table.value, columns=Range('working', 'A1').horizontal.value)


# In[10]:

appended['Site'] = appended['Site (DFA)']
appended['Destination URL'] = appended['Click-through URL']

appended = appended.drop('Site (DFA)', 1)
appended = appended.drop('Click-through URL', 1)


##### Add Week and Video columns

# In[11]:

appended['Week'] = appended['Date'].min()
appended['Video Completions'] = 0
appended['Video Views'] = 0


##### Using the list of actions in the 'Action Reference' tab of the Excel sheet, set lists for each action category.

# In[12]:

a_actions = Range('Action_Reference', 'A2').vertical.value
b_actions = Range('Action_Reference', 'B2').vertical.value
c_actions = Range('Action_Reference', 'C2').vertical.value
d_actions = Range('Action_Reference', 'D2').vertical.value
e_actions = Range('Action_Reference', 'E2').vertical.value

col_head = Range('working', 'A1').horizontal.value


##### Set the actions to lists and search the DataFrame columns for each one, summing each value when found.

# In[13]:

a_actions = list(set(a_actions).intersection(col_head))
b_actions = list(set(b_actions).intersection(col_head))
c_actions = list(set(b_actions).intersection(col_head))
d_actions = list(set(d_actions).intersection(col_head))
e_actions = list(set(e_actions).intersection(col_head))


# In[14]:

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


# In[15]:

view_based = list(set(view_through).intersection(col_head))
click_based = list(set(click_through).intersection(col_head))


##### Add columns to the DataFrame for each action category

# In[16]:

appended['A Actions'] = appended[a_actions].sum(axis=1)
appended['B Actions'] = appended[b_actions].sum(axis=1)
appended['C Actions'] = appended[c_actions].sum(axis=1)
appended['D Actions'] = appended[d_actions].sum(axis=1)
appended['E Actions'] = appended[e_actions].sum(axis=1)

appended['F Actions'] = 0

appended['Post-Click Activity'] = appended[click_based].sum(axis=1)
appended['Post-Impression Activity'] = appended[view_based].sum(axis=1)


##### Store Locator

# In[17]:

store_locator = []
k = iter(store_locator)
for item in col_head:
    locator = re.search('Store Locator', item)
    if locator:
        store_locator.append(item)
        k.next()


# In[18]:

SLV_conversions = list(set(store_locator).intersection(col_head))
appended['Store Locator Visits'] = appended[store_locator].sum(axis=1)


##### Traffic Action Totals

# In[19]:

appended['Awareness Actions'] = appended['A Actions'] + appended['B Actions']
appended['Consideration Actions'] = appended['C Actions'] + appended['D Actions']
appended['Traffic Actions'] = appended['Awareness Actions'] + appended['Consideration Actions']


##### Message Categories

# In[20]:

appended['Creative Field 1'] = appended['Creative Field 1'].str.replace('Creative Type: ', '')

appended['Creative Field 1'] = appended['Creative Field 1'].str.replace('(', '')
appended['Creative Field 1'] = appended['Creative Field 1'].str.replace(')', '')
appended['Creative Field 1'] = appended['Creative Field 1'].str.replace('not set', 'TMO Unique Creative')

appended['Message Bucket'] = appended['Creative Field 1'].str.split('_').str.get(0)

appended['Message Category'] = appended['Creative Field 1'].str.split('_').str.get(1)

appended['Message Offer'] = appended['Creative Field 1'].str.split('_').str.get(2)
appended['Message Offer'].fillna(appended['Creative Groups 2'], inplace=True)


##### F Tag

# In[21]:

appended['F Tag'] = 0
appended['Category'] = 0


##### Strip the embedded URL encoding used by BlueKai to get the actual URL.

# In[22]:

appended['Destination URL'] = appended['Destination URL'].str.replace('http://analytics.bluekai.com/site/', '')
appended['Destination URL'] = appended['Destination URL'].str.replace('15991\?phint', '')
appended['Destination URL'] = appended['Destination URL'].str.replace('http://15991\?phint', '')
appended['Destination URL'] = appended['Destination URL'].str.replace('event%3Dclick&phint', '')
appended['Destination URL'] = appended['Destination URL'].str.replace('aid%3D%eadv!&phint', '')
appended['Destination URL'] = appended['Destination URL'].str.replace('pid%3D%epid!&phint', '')
appended['Destination URL'] = appended['Destination URL'].str.replace('cid%3D%ebuy!&phint', '')
appended['Destination URL'] = appended['Destination URL'].str.replace('crid%3D%ecid!&done', '')
appended['Destination URL'] = appended['Destination URL'].str.replace('pid%3D%25epid!&phint', '')
appended['Destination URL'] = appended['Destination URL'].str.replace('%3Fcmpid%3DWTR_DD_DDRFCBK_RQLMKXRCUQZ1042%26csdids%3D%epid!_%eaid!_%ecid!_%eadv!', '')
appended['Destination URL'] = appended['Destination URL'].str.replace('%3F%26csdids%3DADV_DS_%epid!_%eaid!_%ecid!_%eadv!', '')
appended['Destination URL'] = appended['Destination URL'].str.replace('=', '')
appended['Destination URL'] = appended['Destination URL'].str.replace('%2F', '/')
appended['Destination URL'] = appended['Destination URL'].str.replace('%3A', ':')
appended['Destination URL'] = appended['Destination URL'].str.replace('%23', '#')
appended['Destination URL'] = appended['Destination URL'].apply(lambda x: x.split('.html')[0])
appended['Destination URL'] = appended['Destination URL'].apply(lambda x: x.split('?')[0])


# In[23]:

urls = pd.DataFrame(appended[['Destination URL', 'F Tag']].drop_duplicates())


# In[33]:

url_f_tag = np.where(urls['F Tag'] == 0, urls['Destination URL'], urls['F Tag'])

for url in url_f_tag[0:2]:
    with Browser() as browser:
        page = browser.visit(url)
        f_tag_name = input('Enter F Tag name: ')
        f_tag = np.where(appended['Destination URL'] == url, f_tag_name, 'missing')

appended['F Tag'] = f_tag


# In[123]:

browser = Browser()
page = browser.visit(url)
html = browser.html
soup = BeautifulSoup(html)


# In[209]:

iframes = list(soup.find_all("iframe"))
src = []
floodlights = []
j = iter(floodlights)
i = iter(iframes)
for iframe in iframes:
    src.append(iframe.get('src'))
    i.next()
for fls in src:
    tags = re.search('fls.doubleclick', str(fls))
    if tags:
        floodlights.append(tags)
if len(floodlights) < 1:
    print 'No Floodlight tag found on page'


##### Copy data into pivot tab

# In[24]:

sa_columns = sa.columns.tolist()
cfv_columns = cfv.columns.tolist()


# In[25]:

action_tags = sa_columns[sa_columns.index('Clicks') + 1:]


# In[26]:

metrics = ['Media Cost', 'Impressions', 'Clicks', 'Orders', 'Plans', 'Add-a-Line', 'Activations', 'Devices', 'Services', 'Accessories',
           'Prepaid Plans', 'Store Locator Visits', 'A Actions', 'B Actions', 'C Actions', 'D Actions', 'F Actions', 'Awareness Actions',
           'Consideration Actions', 'Traffic Actions', 'Post-Click Activity', 'Post-Impression Activity']

dimensions = ['Week', 'Date', 'Campaign', 'Site', 'Category', 'Destination URL', 'F Tag', 'Message Bucket', 'Message Category', 'Message Offer',
              'Creative', 'Ad', 'Creative Groups 1', 'Creative Groups 2', 'Creative ID', 'Creative Type', 'Creative Field 1', 'Placement',
              'Placement Cost Structure']


# In[27]:

columns = dimensions + metrics + action_tags
columns = list(itertools.chain(columns))


# In[28]:

appended = appended[columns]


# In[29]:

Range('working', 'A1', index=False).value = appended

