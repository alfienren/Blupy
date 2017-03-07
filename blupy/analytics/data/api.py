from config import configFile
import requests
import pandas as pd
from analytics.data.file_io import DataMethods


class ReportingAPI(configFile):

    def __init__(self):
        super(ReportingAPI, self).__init__()
        self.configs = configFile().load_config()
        self.auth = (self.configs['placed_api']['credentials']['username'],
                     self.configs['placed_api']['credentials']['password'])
        self.url = self.configs['placed_api']['url']
        self.metric = self.configs['placed_api']['csv']

    def placed(self):
        t = self.configs['placed_api']['endpoints']['t-mobile']
        m = self.configs['placed_api']['endpoints']['metro']

        tc = dict(t)
        tc.update(m)

        placed_data = pd.DataFrame()

        for key, value in tc.iteritems():
            req = requests.get(self.url + value + self.metric, auth=self.auth)
            if req.status_code != 404:
                data = pd.read_csv(req.json()['urls'][0], delimiter=',')
                data['endpoint'] = key
                placed_data = placed_data.append(data)

        DataMethods().chunk_df(placed_data, 'Sheet2', 'A1')