from xlwings import Range
import pandas as pd
import re

def cfv():

    # Before function is ran, VBA code will create the necessary tabs in order to process correctly. See the
    # documentation in the VBA modules for more information.
    # Workbook needs to be saved in order to load the data into pandas properly
    # Load the Site Activity and Custom Floodlight Variable data into pandas as DataFrames

    cfv = pd.DataFrame(Range('CFV_Temp', 'A1').table.value, columns = Range('CFV_Temp', 'A1').horizontal.value)
    cfv.drop(0, inplace = True)

    return cfv

def sa():

    sheet = Range('Lookup', 'AA1').value
    sa = pd.read_excel(sheet, 'SA_Temp', index_col=None)

    return sa