from config import configFile
import requests
import pandas as pd


class ReportingAPI(configFile):

    def __init__(self):
        super(ReportingAPI, self).__init__()
        self.configs = configFile().load_config()

    def placed(self):
        auth = (self.configs['placed_api']['credentials']['username'],
                self.configs['placed_api']['credentials']['password'])

        tmo_el, tmo_sl, metro_el, metro_sl = self.configs['placed_api']['endpoints']['tmo_el'], \
                                             self.configs['placed_api']['endpoints']['tmo_sl'], \
                                             self.configs['placed_api']['endpoints']['metro_el'], \
                                             self.configs['placed_api']['endpoints']['metro_sl']

        data_endpoints = [tmo_el, tmo_sl, metro_el, metro_sl]
        endpoint_names = ['TMO_EL', 'TMO_SL', 'Metro_EL', 'Metro_SL']

        placed_data = pd.DataFrame()

        for i in range(0, len(data_endpoints)):
            req = requests.get(i, auth=auth)
            data = pd.read_csv(req.json()['urls'][0], delimiter=',')
            data['endpoint'] = endpoint_names[i]
            placed_data = placed_data.append(data)

        return placed_data
