import numpy as np

def additional_columns(data):

    # The DFA field DBM Cost is more accurate for placements using dynamic bidding. If a placement is not using
    # dynamic bidding, DBM Cost = 0. Therefore, if DBM cost does not equal 0, replace the row's media cost with
    # DBM cost. If DBM Cost = 0, Media Cost stays the same.
    data['Media Cost'] = np.where(data['DBM Cost USD'] != 0, data['DBM Cost USD'], data['Media Cost'])

    # Adjust spend to Net to Client
    data['NTC Media Cost'] = 0

    # DBM Cost column is then removed as it is no longer needed.
    data.drop('DBM Cost USD', 1, inplace=True)

    # Add columns for Video Completions and Views, primarily for compatibility between campaigns that run video and
    # don't run video. Those that don't can keep the columns set to zero, but those that have video can then be
    # adjusted with passback data.
    data['Video Completions'] = 0
    data['Video Views'] = 0

    return data

def order_columns():

    dimensions = ['Week', 'Date', 'Month', 'Quarter', 'Campaign', 'Language', 'Site (DCM)', 'Site', 'Click-through URL',
                  'F Tag', 'Category', 'Category_Adjusted', 'Message Bucket', 'Message Category', 'Creative Bucket',
                  'Creative Theme', 'Creative Type', 'Creative Groups 1', 'Creative ID', 'Ad', 'Creative Groups 2',
                  'Creative Field 1', 'Placement Messaging Type', 'Placement', 'Placement ID',
                  'Placement Cost Structure']

    cfv_floodlight_columns = ['OrderNumber (string)',  'Activity','Floodlight Attribution Type',
                              'Plan (string)', 'Device (string)', 'Service (string)', 'Accessory (string)']

    metrics = ['Media Cost', 'NTC Media Cost', 'Impressions', 'Clicks', 'Orders', 'Plans', 'Add-a-Line',
               'Activations', 'Devices', 'Services', 'Accessories', 'Postpaid Plans', 'Prepaid Plans', 'eGAs',
               'Store Locator Visits', 'A Actions', 'B Actions', 'C Actions', 'D Actions', 'E Actions', 'F Actions',
               'Awareness Actions', 'Consideration Actions', 'Traffic Actions', 'Post-Click Activity',
               'Post-Impression Activity', 'Video Views', 'Video Completions', 'Prepaid GAs', 'Postpaid GAs',
               'Prepaid SIMs', 'Postpaid SIMs', 'Prepaid Mobile Internet', 'Postpaid Mobile Internet',
               'Prepaid Phone', 'Postpaid Phone', 'Total GAs', 'DDR New Devices', 'DDR Add-a-Line']

    new_columns = dimensions + metrics + cfv_floodlight_columns

    return list(new_columns)

