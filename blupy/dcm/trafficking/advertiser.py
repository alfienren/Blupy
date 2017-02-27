import sys

import numpy as np
import pandas as pd
from pandas.io.json import json_normalize
from xlwings import Workbook, Range

from dcm.dcm_api import DCM_API
from analytics.data_refresh.data import DataMethods
from config import configFile


class Advertiser(DCM_API):

    def __init__(self):
        super(Advertiser, self).__init__()

    def advertiser_id(self, tab, cell):
        if Range(tab, cell).value == 'T-Mobile':
            adv_id = self.configs['advertiser_ids']['T-Mobile']
        elif Range(tab, cell).value == 'MetroPCS':
            adv_id = self.configs['advertiser_ids']['MetroPCS']
        else:
            sys.exit('An Advertiser must be selected')

        return adv_id

    def list_campaign_names_ids(self):
        adv_id = Advertiser().advertiser_id(self.FLOODLIGHT_INFO_LIST, 'I1')

        campaigns = self.service.campaigns().list(profileId=self.prof_id, advertiserIds=adv_id,
                                          fields='campaigns(id,name),nextPageToken').execute()

        campaigns_df = json_normalize(campaigns['campaigns'])

        DataMethods().chunk_df(campaigns_df, 'Campaigns', 'A3')

    def placement_traffic_sheet(self):
        camp_ids = pd.DataFrame(Range(self.TRAFFIC_SHEET, 'O1').table.value,
                                columns=Range(self.TRAFFIC_SHEET, 'O1').horizontal.value)
        camp_ids.drop(0, inplace=True)

        adv_id = Advertiser().advertiser_id(self.TRAFFIC_SHEET, 'I1')
        placements_json = self.service.placements()
        request = placements_json.list(profileId=self.prof_id,
                                       advertiserIds=adv_id,
                                       campaignIds=camp_ids['Campaign IDs'].tolist(),
                                       fields='nextPageToken,placements(campaignId,compatibility,' +
                                              'id,name,pricingSchedule(capCostOption,endDate,' +
                                              'pricingPeriods(rateOrCostNanos,units),pricingType,startDate),' +
                                              'siteId,size(height,width))')
        traffic_sheet = pd.DataFrame()

        while request is not None:
            placement = request.execute()
            try:
                placements_norm = json_normalize(placement['placements'], meta='pricingSchedule')
                rate_units = pd.DataFrame.from_records(
                    placements_norm['pricingSchedule.pricingPeriods'].apply(pd.Series)[0])
                placements_norm = pd.concat([placements_norm, rate_units], axis=1)
            except IndexError:
                break

            traffic_sheet = traffic_sheet.append(placements_norm)
            request = placements_json.list_next(request, placement)

        del traffic_sheet['pricingSchedule.pricingPeriods']

        traffic_sheet.rename(columns={'pricingSchedule.capCostOption': 'Cap Cost',
                                      'pricingSchedule.endDate': 'End Date',
                                      'pricingSchedule.startDate': 'Start Date',
                                      'pricingSchedule.pricingType': 'Cost Structure',
                                      'rateOrCostNanos': 'Rate (USD)',
                                      'units': 'Units'}, inplace=True)

        traffic_sheet['Dimensions'] = traffic_sheet['size.width'].astype(str) + 'x' + traffic_sheet['size.height'].astype(
            str)

        traffic_sheet['Rate (USD)'] = traffic_sheet['Rate (USD)'].astype(float) / 1000000000
        traffic_sheet['Cost (USD)'] = np.where(traffic_sheet['Cost Structure'].str.contains('FLAT_RATE') == True,
                                               traffic_sheet['Rate (USD)'],
                                               np.where(traffic_sheet['Cost Structure'].str.contains('CPM') == True,
                                                        traffic_sheet['Units'].astype(float) / 1000 *
                                                        traffic_sheet['Rate (USD)'].astype(float),
                                                        np.where(
                                                            traffic_sheet['Cost Structure'].str.contains('CPC') == True,
                                                            traffic_sheet['Units'].astype(float) *
                                                            traffic_sheet['Rate (USD)'].astype(float), 0)))

        del traffic_sheet['size.height']
        del traffic_sheet['size.width']

        #camp_ids.set_index('Campaign IDs', inplace=True)
        #traffic_sheet.set_index('Campaign ID', inplace=True)

        #traffic_sheet = pd.merge(traffic_sheet, camp_ids, how='left', left_index=True, right_index=True).reset_index()

        DataMethods().chunk_df(traffic_sheet, self.TRAFFIC_SHEET, 'A' + str(len(camp_ids) + 3))