from xlwings import Range
import pandas as pd

def cfv_data():
    # Before function is ran, VBA code will create the necessary tabs in order to process correctly. See the
    # documentation in the VBA modules for more information.
    # Workbook needs to be saved in order to load the data into pandas properly
    # Load the Site Activity and Custom Floodlight Variable data into pandas as DataFrames

    cfv = pd.DataFrame(Range('CFV_Temp', 'A1').table.value, columns = Range('CFV_Temp', 'A1').horizontal.value)
    cfv.drop(0, inplace = True)

    return cfv

def sa_data():

    sa = pd.DataFrame(pd.read_excel(Range('Lookup', 'AA1').value, 'SA_Temp', index_col=None))

    return sa

def columns(sa):

    dimensions = ['Week', 'Date', 'Campaign', 'Site (DCM)', 'Click-through URL', 'F Tag', 'Category', 'Message Bucket',
                  'Message Category', 'Creative Type', 'Creative Groups 1', 'Creative ID', 'Message Offer', 'Creative',
                  'Ad', 'Creative Groups 2', 'Creative Field 1', 'Placement', 'Placement ID',
                  'Placement Cost Structure', 'OrderNumber (string)',  'Activity','Floodlight Attribution Type',
                  'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)']

    if 'DDR' in sa['Campaign'] == True:

        metrics = ['Media Cost', 'NTC Media Cost', 'Impressions', 'Clicks', 'Orders', 'Plans', 'Add-a-Line',
                   'Activations', 'Devices', 'Services', 'Accessories',
                   'Postpaid Plans', 'Prepaid Plans', 'eGAs', 'Store Locator Visits', 'A Actions', 'B Actions', 'C Actions',
                   'D Actions', 'E Actions', 'F Actions', 'Awareness Actions', 'Consideration Actions',
                   'Traffic Actions', 'Post-Click Activity', 'Post-Impression Activity', 'Video Views',
                   'Video Completions', 'Prepaid GAs', 'Postpaid GAs', 'Prepaid SIMs', 'Postpaid SIMs',
                   'Prepaid Mobile Internet', 'Postpaid Mobile Internet', 'Prepaid Phone', 'Postpaid Phone',
                   'DDR New Devices', 'DDR Add-a-Line']

    else:

        metrics = ['Media Cost', 'NTC Media Cost', 'Impressions', 'Clicks', 'Orders', 'Plans', 'Add-a-Line',
                   'Activations', 'Devices', 'Services', 'Accessories', 'Postpaid Plans', 'Prepaid Plans', 'eGAs',
                   'Store Locator Visits', 'A Actions', 'B Actions', 'C Actions', 'D Actions', 'E Actions', 'F Actions',
                   'Awareness Actions', 'Consideration Actions', 'Traffic Actions', 'Post-Click Activity',
                   'Post-Impression Activity', 'Video Views', 'Video Completions']

    sa_columns = list(sa.columns)
    tag_columns = sa_columns[sa_columns.index('DBM Cost USD') + 1:]
    new_columns = dimensions + metrics + tag_columns

    return new_columns