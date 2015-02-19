"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
DDR Custom Floodlight Data Transform
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

import itertools
import pandas as pd
import numpy as np
from xlwings import Workbook, Range, Sheet

def ddr_top_15_devices():

    wb = Workbook.caller()

    cfv = pd.DataFrame(Range('CFV_Temp', 'A1').table.value, columns = Range('CFV_Temp', 'A1').horizontal.value)
    cfv.drop(0, inplace=True)

    sheet = Range('Lookup', 'G1').value

    ddr = pd.DataFrame(pd.read_csv(sheet[:sheet.rindex('\\')] + '\\_\\devices.csv'))

    cfv['Device IDs'] = cfv['Device (string)'].str.split(',')
    ddr_devices = pd.Series(list(np.where((cfv['Campaign'].str.contains('DDR') == True) |
                                          (cfv['Campaign'].str.contains('Q1_Brand Remessaging') == True),
                                          cfv['Device IDs'], np.NaN)))

    ddr_devices.dropna(inplace = True)
    ddr_devices = list(itertools.chain(*ddr_devices))

    while '' in ddr_devices: ddr_devices.remove('')

    device_counts = pd.DataFrame(pd.value_counts(pd.Series(ddr_devices).values, sort = True)[0:15])
    device_counts['Device Name'] = 1
    device_counts['Subcategory'] = 1

    Range('DDR', 'A1').value = ddr
    Range('Summary', 'B1').value = device_counts

    Sheet('Summary').activate()

    i = 0
    for cell in Range('Summary', 'A2:' + 'A' + str(len(device_counts) + 1)):
        i = i + 1
        cell.value = i

    for cell in Range('Summary', 'D2').vertical:
        id = cell.offset(0, -2).get_address(False, False, False)
        cell.formula = '=IFERROR(INDEX(DDR!B:B,MATCH(Summary!' + id + ',DDR!H:H,0)),"na")'

    for cell in Range('Summary', 'E2').vertical:
        id = cell.offset(0, -3).get_address(False, False, False)
        cell.formula = '=IFERROR(INDEX(DDR!G:G,MATCH(Summary!' + id + ',DDR!H:H,0)),"na")'

    for cell in Range('Summary', 'F2').vertical:
        subcategory = cell.offset(0, -1).get_address(False, False, False)
        device = cell.offset(0, -2).get_address(False, False, False)
        cell.formula = '=' + subcategory + '&" - "&' + device