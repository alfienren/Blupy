import json


class configFile(object):

    def __init__(self):
        super(configFile, self).__init__()
        self.config_file = 'config.json'

    def load_config(self):
        with open(self.config_file) as configfile:
            config = json.load(configfile)

        return config

