__author__ = 'aarschle1'

import numpy as np

def floodlight(data):

    # CFV columns for Plans, Services, etc. that were created earlier have blank values replaced with 0.
    data['Plans'].fillna(0, inplace=True)
    data['Services'].fillna(0, inplace=True)
    data['Devices'].fillna(0, inplace=True)
    data['Accessories'].fillna(0, inplace=True)
    data['Orders'].fillna(0, inplace=True)

    # If the count of plans, services, accessories, devices, or orders is less than 1, the string is set to blank. If
    # the count is 1 or greater, the string is associated to the count.
    data['Plan (string)'] = np.where(data['Plans'] < 1, '', data['Plan (string)'])
    data['Service (string)'] = np.where(data['Services'] < 1, '', data['Service (string)'])
    data['Accessory (string)'] = np.where(data['Accessories'] < 1, '', data['Accessory (string)'])
    data['Device (string)'] = np.where(data['Devices'] < 1, "", data['Device (string)'])
    data['OrderNumber (string)'] = np.where(data['Orders'] < 1, '', data['OrderNumber (string)'])
    data['Activity'] = np.where(data['Orders'] < 1, '', data['Activity'])
    data['Floodlight Attribution Type'] = np.where(data['Orders'] < 1, '', data['Floodlight Attribution Type'])
    data['Devices'] = np.where(data['Device (string)'].str.contains('nan') == True, 0, data['Devices'])

    return data