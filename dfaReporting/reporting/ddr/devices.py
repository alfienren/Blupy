import itertools

import pandas as pd
import numpy as np
from xlwings import Range, Sheet


def device_feed():
    path = Range('Action_Reference', 'AE1').value
    device_feed = pd.read_table(path)

    return device_feed


def excluded_devices():
    excluded = Range('Lookup', 'O2').value

    return str(excluded)


def top_15_devices(cfv, feed_path, excl_devices):
    Sheet.add('DDR')
    Sheet.add('Summary')

    device_text_file = pd.read_table(feed_path)

    excluded = excl_devices

    cfv['Device IDs'] = cfv['Device (string)'].str.split(',')

    cfv['Plan Names'] = cfv['Plan (string)'].str.split(',')

    ddr_devices = pd.Series(list(np.where((cfv['Campaign'].str.contains('DDR') == True) |
                                          (cfv['Campaign'].str.contains('Brand Remessaging') == True),
                                          cfv['Device IDs'], np.NaN)))

    ddr_plans = pd.Series(list(np.where((cfv['Campaign'].str.contains('DDR') == True) |
                                        (cfv['Campaign'].str.contains('Brand Remessaging') == True),
                                        cfv['Plan Names'], np.NaN)))

    ddr_devices.dropna(inplace = True)
    ddr_plans.dropna(inplace = True)

    ddr_devices = list(itertools.chain(*ddr_devices))
    ddr_plans = list(itertools.chain(*ddr_plans))

    while '' in ddr_devices: ddr_devices.remove('')
    while '' in ddr_plans: ddr_plans.remove('')
    while excluded in ddr_devices: ddr_devices.remove(excluded)

    device_counts = pd.DataFrame(pd.value_counts(pd.Series(ddr_devices).values, sort = True)[0:15])
    device_counts['Device Name'] = 1

    plan_counts = pd.DataFrame(pd.value_counts(pd.Series(ddr_plans).values, sort = True)[0:15])

    Range('DDR', 'A1', index = False).value = device_text_file

    Range('Summary', 'B1').value = device_counts

    Range('Summary', 'I1').value = plan_counts

    Sheet('Summary').activate()

    # Rank

    i = 0
    for cell in Range('Summary', 'A2:' + 'A' + str(len(device_counts) + 1)):
        i += 1
        cell.value = i

    j = 0
    for cell in Range('Summary', 'H2:' + 'H' + str(len(plan_counts) + 1)):
        j += 1
        cell.value = j

    # Device Name

    for cell in Range('Summary', 'D2').vertical:
        ids = cell.offset(0, -2).get_address(False, False, False)
        cell.formula = '=IFERROR(INDEX(DDR!A:A,MATCH(Summary!' + ids + ',DDR!G:G,0)),"na")'

    Range('Summary', 'A1:C1').value = 'Rank', 'Device SKU', 'Count'
    Range('Summary', 'H1').value = 'Rank'
    Range('Summary', 'I1').value = 'Plan Name'