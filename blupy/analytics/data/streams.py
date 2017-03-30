from dcm.reports import DCMReports
from dcm.dcm_api import DCM_API
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
        cfv = self.run_and_download_report(repId, save_path)

        #cfv = pd.read_excel(save_path)
        cfv = Floodlights().custom_variables(cfv)
        cfv = Floodlights().ddr_custom_variables(cfv)

        DataMethods().chunk_df(cfv, 'Sheet1', 'A1')

    def reference_table(self):
        pass