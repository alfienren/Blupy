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
        self.t = self.configs['placed_api']['endpoints']['t-mobile']
        self.m = self.configs['placed_api']['endpoints']['metro']
        self.tc = dict(self.t)
        self.tc.update(self.m)

    def placed(self):
        placed_data = pd.DataFrame()
        for key, value in self.tc.iteritems():
            req = requests.get(self.url + value + self.metric, auth=self.auth)
            try:
                data = pd.read_csv(req.json()['urls'][0], delimiter=',')
                data['endpoint'] = key
                placed_data = placed_data.append(data)
            except KeyError:
                pass

        placed_data['start_date'] = pd.to_datetime(placed_data['start_date'], unit='ms')
        placed_data['end_date'] = pd.to_datetime(placed_data['end_date'], unit='ms')

        placed_data['unique_id'] = placed_data['name'].astype(str) + '_' + placed_data['type'].astype(str) + \
                                   placed_data['publisher_id'].astype(str) + '_' + \
                                   placed_data['tactic_ids'].astype(str) + '_' + \
                                   placed_data['id'].astype(str) + '_' + \
                                   placed_data['endpoint'].astype(str)

        placed_data['unique_id'].apply(lambda x: str(x).replace('nan', ''))

        DataMethods().chunk_df(placed_data, 'Sheet2', 'A1')

    def get_id_status_codes(self):
        pass