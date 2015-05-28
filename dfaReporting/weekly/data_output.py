import numpy as np
import datetime
import arrow
import pandas as pd

def mondays(dates):

    d = dates.toordinal()
    last_monday = d - 7
    monday = last_monday - (last_monday % 7)
    monday = datetime.date.fromordinal(monday) + datetime.timedelta(1)

    return monday

def columns(sa):

    dimensions = ['Week', 'Date', 'Campaign', 'Site (DCM)', 'Site', 'Click-through URL', 'F Tag', 'Category', 'Category Adjusted',
                  'Message Bucket', 'Message Category', 'Message Offer', 'Creative Type', 'Creative Groups 1',
                  'Creative ID', 'Creative', 'Ad', 'Creative Groups 2', 'Creative Field 1', 'Placement', 'Placement ID',
                  'Placement Cost Structure']

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

def output(data):

    # The DFA field DBM Cost is more accurate for placements using dynamic bidding. If a placement is not using
    # dynamic bidding, DBM Cost = 0. Therefore, if DBM cost does not equal 0, replace the row's media cost with
    # DBM cost. If DBM Cost = 0, Media Cost stays the same.
    data['Media Cost'] = np.where(data['DBM Cost USD'] != 0, data['DBM Cost USD'], data['Media Cost'])

    # Adjust spend to Net to Client
    data['NTC Media Cost'] = 0

    # DBM Cost column is then removed as it is no longer needed.
    data.drop('DBM Cost USD', 1, inplace=True)

    # Create week column by taking the oldest date in the data
    data['Date2'] = pd.to_datetime(data['Date'])
    data['Week'] = data['Date2'].apply(lambda x: mondays(x))
    #data['Month'] = data['Date2'].apply(lambda x: arrow.get(x).format('MMMM'))
    data.drop('Date2', axis = 1, inplace = True)

    # Add columns for Video Completions and Views, primarily for compatibility between campaigns that run video and
    # don't run video. Those that don't can keep the columns set to zero, but those that have video can then be
    # adjusted with passback data.

    data['Video Completions'] = 0
    data['Video Views'] = 0

    return data