
# coding: utf-8

# In[1]:

import pandas as pd
import numpy as np
from xlwings import Workbook, Range
import re
import itertools
import html5lib
import requests
from bs4 import BeautifulSoup
import lxml
import urllib


# In[2]:

wb = Workbook("C:/Users/aarschle1/Google Drive/Optimedia/T-Mobile/Projects/Weekly_Reporting/Opti_DFA_Weekly_Reporting.xlsm")


# In[3]:

sa2 = pd.DataFrame(pd.read_excel(wb.fullname, 'SA_Temp', index_col = None))
#cfv2 = pd.DataFrame(pd.read_excel(wb.fullname, 'CFV_Temp', index_col = None))

cfv2 = pd.DataFrame(Range('CFV_Temp', 'A1').table.value, columns = Range('CFV_Temp', 'A1').horizontal.value)
cfv2.drop(0, inplace=True)


# In[4]:

def chunk_df(df, sheet, startcell, chunk_size = 100):
    if len(df) <= (chunk_size + 1):
        Range(sheet, startcell, index = False, header = True).value = df
    else:
        Range(sheet, startcell, index = False).value = list(df.columns)
        c = re.match(r"([a-z]+)([0-9]+)", startcell[0] + str(int(startcell[1]) + 1), re.I)
        row = c.group(1)
        col = int(c.group(2))
        
        for chunk in (df[rw:rw + chunk_size] for rw in 
                      range(0, len(df), chunk_size)):
            Range(sheet, row + str(col), index = False, header = False).value = chunk
            col += chunk_size


# In[5]:

sa = sa2
cfv = cfv2


# In[8]:

cfv['Orders'] = 1

cfv['Plans'] = np.where(cfv['Plan (string)'] != '', cfv['Plan (string)'].str.count(',') + 1, 0)
cfv['Services'] = np.where(cfv['Service (string)'] != '', cfv['Service (string)'].str.count(',') + 1, 0)
cfv['Accessories'] = np.where(cfv['Accessory (string)'] != '', cfv['Accessory (string)'].str.count(',') + 1, 0)
cfv['Devices'] = np.where(cfv['Device (string)'] != '', cfv['Device (string)'].str.count(',') + 1, 0)
cfv['Add-a-Line'] = cfv['Service (string)'].str.count('ADD')
cfv['Activations'] = cfv['Plans'] + cfv['Add-a-Line']

#cfv['Plans'] = cfv['Plan (string)'].str.count(',') + 1
#cfv['Devices'] = cfv['Device (string)'].str.count(',') + 1
#cfv['Services'] = cfv['Service (string)'].str.count(',') + 1
#cfv['Accessories'] = cfv['Accessory (string)'].str.count(',') + 1

cfv['Postpaid Plans'] = np.where(cfv['Plans'] == cfv['Devices'], cfv['Plans'], pd.concat([cfv['Plans'], cfv['Devices']], axis=1).min(axis=1))
cfv['Prepaid Plans'] = np.where((cfv['Plans'] == 0) & (cfv['Devices'] != 0), 0, cfv['Devices'])

cfv['eGAs'] = np.where(cfv['Floodlight Attribution Type'].str.contains('View-through') == True,
                            (cfv['Device (string)'].str.count(',') + 1) / 2,
                            cfv['Device (string)'].str.count(',') + 1)


# In[9]:

devices = cfv['Device (string)'].str.split(',').apply(pd.Series).stack()


# In[10]:

devices.index = devices.index.droplevel(-1)
devices.name = "Device IDs"


# In[11]:

cfv_new = cfv[cfv.columns[0:17]].join(devices)
cfv_new = cfv.append(cfv_new)


# In[12]:

ddr = pd.DataFrame(pd.read_csv('C:/Users/aarschle1/Google Drive/Optimedia/T-Mobile/Projects/Weekly_Reporting/devices_feed.csv'))


# In[13]:

merged = pd.merge(cfv_new, ddr, how = 'left', left_on = 'Device IDs', right_on = 'Device SKU')


# In[40]:

Range('data', 'A1', index = False).value = merged


### Calculations

# In[20]:

merged['Prepaid GAs'] = np.where((merged['Product Subcategory'].str.contains('Prepaid') == True) & 
                                 (merged['Floodlight Attribution Type'].str.contains('View-through') == True), 0.5,
                                  np.where(merged['Product Subcategory'].str.contains('Prepaid') == True, 1, 0))


# In[21]:

merged['Postpaid GAs'] = np.where((merged['Product Subcategory'].str.contains('Postpaid') == True) &
                                  (merged['Floodlight Attribution Type'].str.contains('View-through') == True), 0.5,
                                  np.where(merged['Product Subcategory'].str.contains('Postpaid') == True, 1, 0))


# In[27]:

merged['Prepaid SIMs'] = np.where((merged['Product Category'].str.contains('SIM card') == True) & 
                                  (merged['Product Subcategory'].str.contains('Prepaid') == True) &
                                  (merged['Floodlight Attribution Type'].str.contains('View-through') == True), 0.5,
                                  np.where((merged['Product Category'].str.contains('SIM card') == True) & 
                                           (merged['Product Subcategory'].str.contains('Prepaid') == True), 1, 0))


# In[29]:

merged['Postpaid SIMs'] = np.where((merged['Product Category'].str.contains('SIM card') == True) & 
                                   (merged['Floodlight Attribution Type'].str.contains('View-through') == True) & 
                                   (merged['Product Subcategory'].str.contains('Postpaid') == True), 0.5,
                                   np.where((merged['Product Category'].str.contains('SIM card') == True) & 
                                            (merged['Product Subcategory'].str.contains('Postpaid') == True), 1, 0))


# In[30]:

merged['Prepaid Mobile Internet'] = np.where((merged['Product Category'].str.contains('Mobile Internet') == True) & 
                                             (merged['Product Subcategory'].str.contains('Prepaid') == True) & 
                                             (merged['Floodlight Attribution Type'].str.contains('View-through') == True), 0.5,
                                             np.where((merged['Product Category'].str.contains('Mobile Internet') == True) & 
                                                      (merged['Product Subcategory'].str.contains('Prepaid') == True), 1, 0))


# In[32]:

merged['Postpaid Mobile Internet'] = np.where((merged['Product Category'].str.contains('Mobile Internet') == True) & 
                                              (merged['Product Subcategory'].str.contains('Postpaid') == True) & 
                                              (merged['Floodlight Attribution Type'].str.contains('View-through') == True), 0.5,
                                              np.where((merged['Product Category'].str.contains('Mobile Internet') == True) & 
                                                       (merged['Product Subcategory'].str.contains('Postpaid') == True), 1, 0))


# In[33]:

merged['Prepaid Phone'] = np.where((merged['Product Category'].str.contains('Smartphone') == True) & 
                                   (merged['Product Subcategory'].str.contains('Prepaid') == True) & 
                                   (merged['Floodlight Attribution Type'].str.contains('View-through') == True), 0.5, 
                                   np.where((merged['Product Category'].str.contains('Smartphone') == True) & 
                                            (merged['Product Subcategory'].str.contains('Prepaid') == True), 1, 0))


# In[34]:

merged['Postpaid Phone'] = np.where((merged['Product Category'].str.contains('Smartphone') == True) & 
                                    (merged['Product Subcategory'].str.contains('Postpaid') == True) & 
                                    (merged['Floodlight Attribution Type'].str.contains('View-through') == True), 0.5, 
                                    np.where((merged['Product Category'].str.contains('Smartphone') == True) & 
                                             (merged['Product Subcategory'].str.contains('Postpaid') == True), 1, 0))


# In[37]:

merged['DDR New Devices'] = np.where((merged['Device IDs'].notnull() == True) & 
                                 (merged['Activity'].str.contains('New TMO Order') == True) &
                                 (merged['Floodlight Attribution Type'].str.contains('View-through') == True), 0.5,
                                 np.where((merged['Device IDs'].notnull() == True) & 
                                          (merged['Activity'].str.contains('New TMO Order') == True), 1, 0))


# In[39]:

merged['DDR Add-a-Line'] = np.where((merged['Device IDs'].notnull() == True) & 
                                    (merged['Activity'].str.contains('New My.TMO Order') == True) & 
                                    (merged['Floodlight Attribution Type'].str.contains('View-through') == True), 0.5,
                                     np.where((merged['Device IDs'].notnull() == True) &
                                              (merged['Activity'].str.contains('New My.TMO Order') == True), 1, 0))


### Top 15 Device Count

# In[77]:

ddr_devices = pd.Series(list(np.where(cfv['Campaign'].str.contains('DDR') == True, cfv['Device IDs'], np.NaN)))
ddr_devices.dropna(inplace = True)
ddr_devices = list(itertools.chain(*ddr_devices))
while '' in ddr_devices: ddr_devices.remove('')
pd.value_counts(pd.Series(ddr_devices).values, sort = True)[0:15]

