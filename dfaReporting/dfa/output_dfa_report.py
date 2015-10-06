import numpy as np

def columns():

    dimensions = ['Week', 'Date', 'Month', 'Quarter', 'Campaign', 'Language', 'Site (DCM)', 'Site', 'Click-through URL',
                  'F Tag', 'Category', 'Message Bucket', 'Message Category', 'Creative Bucket', 'Creative Theme',
                  'Creative Type', 'Creative Groups 1', 'Creative ID', 'Creative', 'Ad', 'Creative Groups 2',
                  'Creative Field 1', 'Placement', 'Placement ID', 'Placement Cost Structure']

    cfv_floodlight_columns = ['OrderNumber (string)',  'Activity','Floodlight Attribution Type',
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

        metrics = ['Media Cost', 'NTC Media Cost', 'Impressions', 'Clicks', 'Video Views', 'Video Completions',
                   'Orders', 'Plans', 'Add-a-Line', 'Activations', 'Devices', 'Services', 'Accessories',
                   'Postpaid Plans', 'Prepaid Plans', 'eGAs', 'Store Locator Visits', 'A Actions', 'B Actions',
                   'C Actions', 'D Actions', 'E Actions', 'F Actions', 'Awareness Actions', 'Consideration Actions',
                   'Traffic Actions', 'Post-Click Activity', 'Post-Impression Activity']

    sa_columns = list(sa.columns)
    tag_columns = sa_columns[sa_columns.index('DBM Cost USD') + 1:]
    new_columns = dimensions + metrics + tag_columns + cfv_floodlight_columns

    return new_columns

