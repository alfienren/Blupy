
# coding: utf-8

# In[7]:

import pandas as pd
import numpy as np
from xlwings import Workbook, Range, Sheet


# In[3]:

wb = Workbook("C:/Users/aarschle1/Google Drive/Optimedia/T-Mobile/Projects/Weekly_Reporting/Opti_DFA_Weekly_Reporting.xlsm")


# In[148]:

sa = Range("CFV_Temp", "A1").table.value
cfv = Range("SA_Temp", "A1").table.value


# In[149]:

sa = pd.DataFrame(sa)
cfv = pd.DataFrame(cfv)


# In[167]:

sa.head(5)


# In[169]:

merged = pd.merge(cfv, sa, how='outer', on=1)


# In[165]:

Range('working', 'A1').value = merged


# In[ ]:



