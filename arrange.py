
# coding: utf-8

# In[1]:

import pandas as pd
import numpy as np
from xlwings import Workbook, Range, Sheet


# In[2]:

wb = Workbook("C:/Users/aarschle1/Google Drive/Optimedia/T-Mobile/Projects/Weekly_Reporting/Opti_DFA_Weekly_Reporting.xlsm")


# In[215]:

sa = Range("SA_Temp", "A2").table.value
cfv = Range("CFV_Temp", "A2").table.value


# In[227]:

sa = pd.DataFrame(sa, columns = Range("SA_Temp", "A1").horizontal.value)
cfv = pd.DataFrame(cfv, columns = Range("CFV_Temp", "A1").horizontal.value)


# In[228]:

appended = sa.append(cfv)


# In[229]:

Range('working', 'A1').value = appended


# In[230]:

appended = appended.groupby(['Campaign', 'Date', 'Site', 'Creative', 'Click-through URL', 'Ad', 'Creative Groups 1',
                             'Creative Groups 2', 'Creative ID', 'Creative Type', 'Creative Field 1', 'Placement',   
                             'Placement Cost Structure', 'Device (string)', 'Floodlight Attribution Type', 'OrderNumber (string)',
                             'Plan (string)', 'Service (string)']).aggregate(np.sum)


# In[231]:

appended = appended.fillna(0)


# In[222]:

Range('working', 'A1').value = appended
appended = Range('working', 'A2').table.value
appended = pd.DataFrame(appended, columns = Range('working', 'A1').horizontal.value)


# In[223]:

view_based = Range('Action_Reference', 'A2').vertical.value
click_based = Range('Action_Reference', 'B2').vertical.value
a_actions = Range('Action_Reference', 'C2').vertical.value
b_actions = Range('Action_Reference', 'D2').vertical.value
c_actions = Range('Action_Reference', 'E2').vertical.value
d_actions = Range('Action_Reference', 'F2').vertical.value
e_actions = Range('Action_Reference', 'G2').vertical.value

col_head = Range('working', 'A1').horizontal.value


# In[224]:

a_actions = list(set(a_actions).intersection(col_head))
b_actions = list(set(b_actions).intersection(col_head))
c_actions = list(set(b_actions).intersection(col_head))
d_actions = list(set(d_actions).intersection(col_head))
e_actions = list(set(e_actions).intersection(col_head))

view_based = list(set(view_based).intersection(col_head))
click_based = list(set(click_based).intersection(col_head))


# In[225]:

appended['A Actions'] = appended[a_actions].sum(axis=1)
appended['B Actions'] = appended[b_actions].sum(axis=1)
appended['C Actions'] = appended[c_actions].sum(axis=1)
appended['D Actions'] = appended[d_actions].sum(axis=1)
appended['E Actions'] = appended[e_actions].sum(axis=1)

appended['Click Based'] = appended[click_based].sum(axis=1)
appended['View Based'] = appended[view_based].sum(axis=1)


# In[226]:

Range('Sheet1', 'A1').value = appended


# In[ ]:



