import getpass

import httplib2
from oauth2client.file import Storage
from oauth2client.client import flow_from_clientsecrets
from oauth2client.tools import run_flow, argparser
from apiclient.discovery import build

from xlwings import Workbook


class DCM_API(object):

    def __init__(self):
        Workbook.caller()
        self.API_NAME = 'dfareporting'
        self.API_VERSION = 'v2.7'
        self.API_SCOPES = ['https://www.googleapis.com/auth/dfatrafficking',
                      'https://www.googleapis.com/auth/dfareporting']
        self.service = self.authenticate_user()
        self.prof_id = self.profile_id(self.service)
        self.fl = self.service.floodlightActivities()
        self.CREATE_FLOODLIGHTS = 'Create_Floodlights'
        self.FLOODLIGHT_INFO_LIST = 'Get_Floodlights'
        self.GENERATE_FLOODLIGHT_TAGS = 'Generate_Tags'
        self.UPDATE_FLOODLIGHTS = 'Update_Floodlights'
        self.PIGGYBACK_PIXELS = 'Implement_Pixels'
        self.TRAFFIC_SHEET = 'Placement_Traffic'
        self.CAMPAIGN_SHEET = 'Campaigns'
        self.LIST_FLOODLIGHT_PIXELS = 'Get_Pixels'
        self.SITEMAP = 'Get_Sitemap'

    def authenticate_user(self):
        storage_path = r'S:\SEA-Media\Analytics\T-Mobile\_\api_credentials\user_credentials\saved_user_creds_'
        user_name = getpass.getuser()

        storage = Storage(storage_path + user_name + '.dat')

        credentials = storage.get()

        if credentials is None or credentials.invalid:
            credentials = run_flow(
                flow_from_clientsecrets(r'S:\SEA-Media\Analytics\T-Mobile\_\api_credentials\client_secret.json',
                                        scope=self.API_SCOPES),
                storage, argparser.parse_args([]))

        auth = credentials.authorize(httplib2.Http())

        service = build(self.API_NAME, self.API_VERSION, http=auth)

        return service

    @staticmethod
    def profile_id(service):
        prof_id = service.userProfiles().list().execute()['items'][0]['profileId']

        return prof_id