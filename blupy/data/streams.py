from dcm.reports import DCMReports
from floodlights import Floodlights
from file_io import DataMethods


class DatoramaStreams(DCMReports):
    def __init__(self):
        super(DatoramaStreams, self).__init__()

    def schedule(self):
        pass

    def custom_floodlights(self):
        save_path = r'C:/Users/aarschle1/Desktop/cfv2.xlsx'
        repId = self.configs['report_ids']['datoramaCFV']
        cfv = self.run_and_download_report(repId)

        cfv = Floodlights().custom_variables(cfv)
        del cfv['Device_reg']

        DataMethods().chunk_df(cfv, 'Sheet1', 'A1')

    def reference_table(self):
        repId = self.configs['report_ids']['datoramaReferenceTable']
        ref = self.run_and_download_report(repId)