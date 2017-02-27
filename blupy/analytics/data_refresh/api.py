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

        base_tmo_el, base_tmo_sl, base_metro_el, base_metro_sl = self.configs['placed_api']['endpoints']["base"][
                                                                     'tmo_el'], \
                                                                 self.configs['placed_api']['endpoints']["base"][
                                                                     'tmo_sl'], \
                                                                 self.configs['placed_api']['endpoints']["base"][
                                                                     'metro_el'], \
                                                                 self.configs['placed_api']['endpoints']["base"][
                                                                     'metro_sl']

        prospect_tmo_el, prospect_tmo_sl, prospect_metro_el, prospect_metro_sl = \
        self.configs['placed_api']['endpoints']["prospect"][
            'tmo_el'], \
        self.configs['placed_api']['endpoints']["prospect"][
            'tmo_sl'], \
        self.configs['placed_api']['endpoints']["prospect"][
            'metro_el'], \
        self.configs['placed_api']['endpoints']["prospect"][
            'metro_sl']

        base_data_endpoints = [base_tmo_el, base_tmo_sl, base_metro_el, base_metro_sl]
        prospect_data_endpoints = [prospect_tmo_el, prospect_tmo_sl, prospect_metro_el, prospect_metro_sl]
        data_endpoints = base_data_endpoints + prospect_data_endpoints
        endpoint_names = ['Base_TMO_EL', 'Base_TMO_SL', 'Base_Metro_EL', 'Base_Metro_SL',
                          'Prospect_TMO_EL', 'Prospect_TMO_SL', 'Prospect_Metro_EL', 'Prospect_Metro_SL']

        placed_data = pd.DataFrame()

        for i in range(0, len(data_endpoints)):
            req = requests.get(i, auth=auth)
            data = pd.read_csv(req.json()['urls'][0], delimiter=',')
            data['endpoint'] = endpoint_names[i]
            placed_data = placed_data.append(data)

        return placed_data
