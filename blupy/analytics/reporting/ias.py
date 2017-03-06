from xlwings import Sheet, Range, Workbook
import string
import pandas as pd
import numpy as np
from analytics.data.io import DataMethods


class IASReporting(DataMethods):

    def __init__(self):
        super(IASReporting, self).__init__()

    def download_data(self):
        data_sheet = 'Sheet1'
        data = pd.DataFrame()
        wb2 = Workbook(Range(data_sheet, 'WWW1').value)
        wb2.set_current()

        sheets = Sheet.all()[:-1]
        for i in sheets:
            c = Range(i, 'A50').offset(0, 1)
            cell_range = Range(Range.get_address(c.current_region))
            col = string.uppercase[cell_range.column - 1]
            row = str(cell_range.row)
            cell = 'B' + row

            d = pd.DataFrame(Range(i, cell).table.value,
                              columns=Range(i, cell).horizontal.value)
            d.drop(0, inplace=True)
            d.fillna(0, inplace=True)
            d['Date'] = pd.to_datetime(d['Date'])

            if 'External Placement ID' in d.columns:
                del d['External Placement ID']
            if 'AdServer Placement ID' in d.columns:
                del d['AdServer Placement ID']
            if 'Publisher Name' in d.columns:
                d.rename(columns={'Publisher Name': 'Media Partner Name'}, inplace=True)

            if i.name == 'Firewall Activity':
                block_status = d[0:6]
                block_status.drop_duplicates(inplace=True)
                block_status.set_index(['Campaign Name', 'Media Partner Name', 'Placement Name'], inplace=True)
                block_status.drop([col for col in block_status.columns if 'Blocking Status' not in col], axis=1,
                                  inplace=True)
                d.drop([col for col in d.columns if 'Blocking Status' in col], axis=1, inplace=True)

            if i.name == 'Traffic by Country':
                d.drop(d.columns[d.sum() > 100000], axis=1, inplace=True)

            if data.empty == True:
                data = data.append(d)
            else:
                #data.set_index(['Date', 'Campaign Name', 'Media Partner Name', 'Placement Name'], inplace=True)
                #d.set_index(['Date', 'Campaign Name', 'Media Partner Name', 'Placement Name'], inplace=True)
                #data = pd.merge(data, d, how='left', left_index=True, right_index=True).reset_index()
                data = data.append(d)

        data = pd.pivot_table(data, index=['Date', 'Campaign Name', 'Media Partner Name', 'Placement Name'],
                              aggfunc=np.sum).reset_index()

        if block_status.empty != True:
            data.set_index(['Campaign Name', 'Media Partner Name', 'Placement Name'], inplace=True)
            data = pd.merge(data, block_status, how='left', right_index=True, left_index=True).reset_index()

        self.wb.set_current()

        if Range(data_sheet, 'A1').value is None:
            DataMethods().chunk_df(data, data_sheet, 'A1')
        else:
            dat = pd.read_excel(self.wb.fullname, data_sheet, index_col=None)
            data = dat.append(data)
            data.drop_duplicates(inplace=True)
            DataMethods().chunk_df(data, data_sheet, 'A1')

        wb2.close()
