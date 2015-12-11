import numpy as np
import pandas as pd
from xlwings import Range

def cfv_data(cfv):

    cfv['Orders'] = 1  # Create orders column in cfv data. Each OrderNumber counts as 1 order

    # Count the number of plans in the Plans column
    cfv['Plans'] = np.where(cfv['Plan (string)'] != np.NaN, cfv['Plan (string)'].str.count(',') + 1, 0)
    # Count number of services in the Service column
    cfv['Services'] = np.where(cfv['Service (string)'] != np.NaN, cfv['Service (string)'].str.count(',') + 1, 0)
    # Count number of Accessories in the Accessories column
    cfv['Accessories'] = np.where(cfv['Accessory (string)'] != np.NaN, cfv['Accessory (string)'].str.count(',') + 1, 0)
    # Count number of devices in the Plans column
    cfv['Devices'] = np.where(cfv['Device (string)'] != np.NaN, cfv['Device (string)'].str.count(',') + 1, 0)
    # Count number of Add-a-Lines in the Service column
    cfv['Add-a-Line'] = cfv['Service (string)'].str.count('ADD')
    # Activations are defined as the sum of Plans and Add-a-Line
    cfv['Activations'] = cfv['Plans'] + cfv['Add-a-Line']

    # Postpaid plans are defined as a Plan + Device. By row, if number of plans is equal to number of devices, Postpaid
    # plans = number of plans. If plans and devices are not equal, use the minimum number.
    cfv['Postpaid Plans'] = abs(np.where(cfv['Plans'] == cfv['Devices'], cfv['Plans'],
                                     pd.concat([cfv['Plans'], cfv['Devices']], axis=1).min(axis=1)))

    # Prepaid plans are defined as the number of Devices with no service and plan. If number of plans and services are
    # 0, count of devices is prepaid. If service and plan are not equal, prepaid plans = 0.
    cfv['Prepaid Plans'] = abs(np.where((cfv['Plans'] == 0) & (cfv['Services'] == 0), cfv['Devices'],
                                    np.where((cfv['Devices'] > cfv['Plans']) & (cfv['Devices'] > cfv['Services']),
                                             cfv['Devices'] - pd.concat([cfv['Plans'], cfv['Services']], axis=1).max(
                                                 axis=1), 0)))

    # The DDR campaign counts view-through order credit at 50%. If the campaign name contains 'DDR' and the Floodlight
    # Attribution Type is View-through, the order is multiplied by 0.5.
    cfv['Orders'] = np.where(
        ((cfv['Campaign'].str.contains('DDR') == True) | (cfv['Campaign'].str.contains('Brand Remessaging') == True)) &
        (cfv['Floodlight Attribution Type'].str.contains('View-through') == True),
        cfv['Orders'] * 0.5, cfv['Orders'])

    # Estimated Gross Adds are calculated as the count of Devices with 50% view-through credit.
    # If Floodlight Attribution Type is equal to View-through, the count of Devices is multiplied by 0.5
    cfv['eGAs'] = np.where(cfv['Floodlight Attribution Type'].str.contains('View-through') == True,
                           (cfv['Device (string)'].str.count(',') + 1) / 2,
                           cfv['Device (string)'].str.count(',') + 1)

    return cfv

def ddr(cfv):

    devices = cfv['Device (string)'].str.split(',').apply(pd.Series).stack()
    devices.index = devices.index.droplevel(-1)
    devices.name = "Device IDs"

    device_cfv = cfv[cfv.columns[0:17]].join(devices)
    cfv = cfv.append(device_cfv)

    device_feed_path = Range('Action_Reference', 'AE1').value

    device_lookup = pd.read_table(device_feed_path)
    excluded_devices = str(Range('Lookup', 'S2').value)
    cfv = pd.merge(cfv, device_lookup, how = 'left', left_on = 'Device IDs', right_on = 'Device SKU')

    cfv['Prepaid GAs'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                   (cfv['Device IDs'].notnull() == True) & (
                                   (cfv['Product Subcategory'].str.contains('Prepaid') == True) | (
                                   cfv['Device IDs'].notnull()) & (cfv['Product Subcategory'].isnull())) &
                                   (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5,
                                  np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) & (
                                  cfv['Device IDs'].notnull() == True) &
                                            (cfv['Product Subcategory'].str.contains('Prepaid') == True) &
                                            (cfv['Floodlight Attribution Type'].str.contains('Click-through') == True)),
                                           1, 0))

    cfv['Postpaid GAs'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                    (cfv['Device IDs'].notnull() == True) & (
                                    cfv['Product Subcategory'].str.contains('Postpaid') == True) &
                                    (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5,
                                   np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                             (cfv['Device IDs'].notnull() == True) & (
                                             cfv['Product Subcategory'].str.contains('Postpaid') == True)), 1, 0))

    cfv['Total GAs'] = cfv['Postpaid GAs'] + cfv['Prepaid GAs']

    cfv['Prepaid SIMs'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                    (cfv['Device IDs'].notnull() == True) & (
                                    cfv['Product Category'].str.contains('SIM card') == True) &
                                    (cfv['Product Subcategory'].str.contains('Prepaid') == True) &
                                    (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5,
                                   np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                             (cfv['Device IDs'].notnull() == True) & (
                                             cfv['Product Category'].str.contains('SIM card') == True) &
                                             (cfv['Product Subcategory'].str.contains('Prepaid') == True) &
                                             (
                                             cfv['Floodlight Attribution Type'].str.contains('Click-through') == True)),
                                            1, 0))

    cfv['Postpaid SIMs'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                     (cfv['Device IDs'].notnull() == True) & (
                                     cfv['Product Category'].str.contains('SIM card') == True) &
                                     (cfv['Floodlight Attribution Type'].str.contains('View-through') == True) &
                                     (cfv['Product Subcategory'].str.contains('Postpaid') == True)), 0.5,
                                    np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                              (cfv['Device IDs'].notnull() == True) & (
                                              cfv['Product Category'].str.contains('SIM card') == True) &
                                              (cfv['Product Subcategory'].str.contains('Postpaid') == True) &
                                              (cfv['Floodlight Attribution Type'].str.contains(
                                                  'Click-through') == True)), 1, 0))

    cfv['Prepaid Mobile Internet'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                               (cfv['Device IDs'].notnull() == True) & (
                                               cfv['Product Category'].str.contains('Mobile Internet') == True) &
                                               (cfv['Product Subcategory'].str.contains('Prepaid') == True) &
                                               (cfv['Floodlight Attribution Type'].str.contains(
                                                   'View-through') == True)), 0.5,
                                              np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                                        (cfv['Device IDs'].notnull() == True) & (
                                                        cfv['Product Category'].str.contains(
                                                            'Mobile Internet') == True) &
                                                        (cfv['Product Subcategory'].str.contains('Prepaid') == True)),
                                                       1, 0))

    cfv['Postpaid Mobile Internet'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                                (cfv['Device IDs'].notnull() == True) & (
                                                cfv['Product Category'].str.contains('Mobile Internet') == True) &
                                                (cfv['Product Subcategory'].str.contains('Postpaid') == True) &
                                                (cfv['Floodlight Attribution Type'].str.contains(
                                                    'View-through') == True)), 0.5,
                                               np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                                         (cfv['Device IDs'].notnull() == True) & (
                                                         cfv['Product Category'].str.contains(
                                                             'Mobile Internet') == True) &
                                                         (cfv['Product Subcategory'].str.contains('Postpaid') == True) &
                                                         (cfv['Floodlight Attribution Type'].str.contains(
                                                             'Click-through') == True)), 1, 0))

    cfv['Prepaid Phone'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                     (cfv['Device IDs'].notnull() == True) & (
                                     cfv['Product Category'].str.contains('Smartphone') == True) &
                                     (cfv['Product Subcategory'].str.contains('Prepaid') == True) &
                                     (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5,
                                    np.where((((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                               (cfv['Device IDs'].notnull() == True) & (
                                               cfv['Product Category'].str.contains('Smartphone') == True) &
                                               (cfv['Product Subcategory'].str.contains('Prepaid') == True) &
                                               (cfv['Floodlight Attribution Type'].str.contains(
                                                   'Click-through') == True))), 1, 0))

    cfv['Postpaid Phone'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                      (cfv['Device IDs'].notnull() == True) & (
                                      cfv['Product Category'].str.contains('Smartphone') == True) &
                                      (cfv['Product Subcategory'].str.contains('Postpaid') == True) &
                                      (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5,
                                     np.where(((cfv['Device IDs'].notnull() == True) & (
                                     cfv['Product Category'].str.contains('Smartphone') == True) &
                                               (cfv['Product Subcategory'].str.contains('Postpaid') == True) &
                                               (cfv['Floodlight Attribution Type'].str.contains(
                                                   'Click-through') == True)), 1, 0))

    cfv['DDR New Devices'] = np.where(
        ((cfv['Device IDs'].str.contains(excluded_devices) == False) & (cfv['Device IDs'].notnull() == True) &
         (cfv['Activity'].str.contains('New TMO Order') == True) &
         (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5,
        np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) & (cfv['Device IDs'].notnull() == True) &
                  (cfv['Activity'].str.contains('New TMO Order') == True) &
                  (cfv['Floodlight Attribution Type'].str.contains('Click-through') == True)), 1, 0))

    cfv['DDR Add-a-Line'] = np.where(
        ((cfv['Device IDs'].str.contains(excluded_devices) == False) & (cfv['Device IDs'].notnull() == True) &
         (cfv['Activity'].str.contains('New My.TMO Order') == True) &
         (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)), 0.5,
        np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) & (cfv['Device IDs'].notnull() == True) &
                  (cfv['Activity'].str.contains('New My.TMO Order') == True) &
                  (cfv['Floodlight Attribution Type'].str.contains('Click-through') == True)), 1, 0))

    return cfv

def clean_cfv(cfv):

    cfv = cfv_data(cfv)
    cfv = ddr(cfv)

    return cfv