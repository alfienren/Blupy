from dcm.dcm_api import DCM_API
import requests
from StringIO import StringIO
import pandas as pd
import numpy as np
import time
from analytics.data.file_io import DataMethods


class DCMReports(DCM_API):

    def __init__(self):
        super(DCMReports, self).__init__()

    def get_report(self, report_id):
        r = self.reporting.get(profileId=self.prof_id, reportId=report_id)
        report = r.execute()

        return report

    def run_report(self, report_id):
        r = self.service.reports().run(profileId=self.prof_id, reportId=report_id).execute()

        return r

    def download_file(self, report_id, file_id, save_path=None):
        f = self.service.reports().files().get(profileId=self.prof_id, reportId=report_id, fileId=file_id).execute()

        while f['status'] == 'PROCESSING':
            time.sleep(10)
            f = self.service.reports().files().get(profileId=self.prof_id, reportId=report_id, fileId=file_id).execute()
            if f['status'] == 'REPORT_AVAILABLE':
                break

        f_loc = self.auth.request(f['urls']['apiUrl'])
        f_loc = f_loc[0]['content-location']

        r = requests.get(f_loc)
        s = r.content[r.content.rindex('Date'):]

        d = pd.read_csv(StringIO(s), na_values=[], keep_default_na=False)

        d['Device (string)'] = d['Device (string)'].astype(str)
        d['Device (string)'] = np.where(d['Device (string)'].str[:6] == '000000',
                                        d['Device (string)'].str[6:],
                                        d['Device (string)'])

        if save_path is not None:
            d.to_excel(save_path)

        return d

    def run_and_download_report(self, report_id, save_path=None):
        f = self.run_report(report_id)
        file_id = f['id']
        dl_report = self.download_file(report_id, file_id, save_path)

        return dl_report