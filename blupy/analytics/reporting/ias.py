from xlwings import Sheet, Range, Workbook
import string
import pandas as pd
from analytics.data_refresh.data import DataMethods


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
            c = Range(i, 'A100').offset(0, 1)
            cell_range = Range(Range.get_address(c.current_region))
            col = string.uppercase[cell_range.column - 1]
            row = str(cell_range.row)
            cell = 'B' + row

            d = pd.DataFrame(Range(i, cell).table.value,
                              columns=Range(i, cell).horizontal.value)
            d.drop(0, inplace=True)

            if 'External Placement ID' in d.columns:
                del d['External Placement ID']
            if 'AdServer Placement ID' in d.columns:
                del d['AdServer Placement ID']
            if 'Publisher Name' in d.columns:
                d.rename(columns={'Publisher Name':'Media Partner Name'}, inplace=True)

            if data.empty == True:
                data = data.append(d)
            else:
                data.set_index(['Date', 'Campaign Name', 'Media Partner Name', 'Placement Name'], inplace=True)
                d.set_index(['Date', 'Campaign Name', 'Media Partner Name', 'Placement Name'], inplace=True)
                data = pd.merge(data, d, how='left', left_index=True, right_index=True).reset_index()

        self.wb.set_current()

        if Range(data_sheet, 'A1').value is None:
            DataMethods().chunk_df(data, data_sheet, 'A1')
        else:
            dat = pd.read_excel(self.wb.fullname, data_sheet, index_col=None)
            data = dat.append(data)
            data.drop_duplicates(inplace=True)
            DataMethods().chunk_df(data, data_sheet, 'A1')

        wb2.close()
