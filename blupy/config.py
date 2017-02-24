import json
import os


class configFile(object):

    def __init__(self):
        #super(configFile, self).__init__()
        self.config_file = os.path.join(os.path.dirname(__file__), 'config.json')
        self.configs = self.load_config()

    def load_config(self):
        with open(self.config_file) as configfile:
            config = json.load(configfile)

        return config

