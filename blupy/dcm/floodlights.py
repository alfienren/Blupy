import sys

import httplib2
import numpy as np
import pandas as pd
from pandas.io.json import json_normalize
from xlwings import Range

from dcm.dcm_api import DCM_API
from data.file_io import DataMethods


class Floodlights(DCM_API):

    def __init__(self):
        super(Floodlights, self).__init__()

    def insert(self):
        tags = pd.DataFrame(Range(self.CREATE_FLOODLIGHTS, 'A1').table.value,
                            columns=Range(self.CREATE_FLOODLIGHTS, 'A1').horizontal.value)

        tags.drop(0, inplace=True)

        tags['Secure'] = np.where(tags['Secure'] == 'Yes', 'True', 'False')

        j = 2
        for index, row in tags.iterrows():
            new_floodlight = {
                'kind': 'dfareporting#floodlightActivity',
                'countingMethod': 'STANDARD_COUNTING',
                'sslRequired': row['Secure'],
                'floodlightActivityGroupId': row['Floodlight Activity Group ID'],
                'name': row['Floodlight Name'],
                'expectedUrl': row['Expected URL']
            }

            self.fl.insert(profileId=self.prof_id, body=new_floodlight).execute()
            Range(self.CREATE_FLOODLIGHTS, 'E' + str(j)).value = 'Created'
            j += 1

    def get(self):
        if Range(self.FLOODLIGHT_INFO_LIST, 'I1').value == 'T-Mobile':
            adv_id = self.configs['advertiser_ids']['T-Mobile']
        elif Range(self.FLOODLIGHT_INFO_LIST, 'I1').value == 'MetroPCS':
            adv_id = self.configs['advertiser_ids']['MetroPCS']
        else:
            sys.exit('An Advertiser must be selected')

        floodlights = self.fl
        request = floodlights.list(profileId=self.prof_id, advertiserId=adv_id)
        floodlight_tags = pd.DataFrame()

        while request is not None:
            flights = request.execute()
            try:
                fl_norm = json_normalize(flights['floodlightActivities'])
            except IndexError:
                break

            floodlight_tags = floodlight_tags.append(fl_norm)
            request = floodlights.list_next(request, flights)

        default_tags = []

        for index, row in floodlight_tags.iterrows():
            if pd.isnull(row)[8] == False:
                default_tags.append('Name: ' + row[8][0]['name'] + ' Tag: ' + row[8][0]['tag'])
            else:
                default_tags.append(None)

        floodlight_tags['defaultTags'] = default_tags

        DataMethods().chunk_df(floodlight_tags, self.FLOODLIGHT_INFO_LIST, 'A3')

    def update(self):
        tags = pd.DataFrame(Range(self.UPDATE_FLOODLIGHTS, 'A1').table.value,
                            columns=Range(self.UPDATE_FLOODLIGHTS, 'A1').horizontal.value)

        tags.drop(0, inplace=True)

        tags['Expected URL'] = np.where(tags['Expected URL'].str.contains('http') == False,
                                        'http://' + tags['Expected URL'] + '.com',
                                        tags['Expected URL'])

        tags['Floodlight Name'] = tags['Floodlight Name'].apply(lambda x: str(x).replace("'", ''))

        batch_floodlights = self.service.new_batch_http_request()

        for index, row in tags.iterrows():
            if row[2] is not None:
                patch_body = {
                    'name': row[1],
                    'expectedUrl': row[2]
                }
            else:
                patch_body = {
                    'name': row[1]
                }
            batch_floodlights.add(self.fl.patch(profileId=self.prof_id, id=row[0], body=patch_body))

        batch_floodlights.execute()

    def generate_all_tags(self):
        if Range(DCM_API().GENERATE_FLOODLIGHT_TAGS, 'I1').value == 'T-Mobile':
            adv_id = self.configs['advertiser_ids']['T-Mobile']
        elif Range(DCM_API().GENERATE_FLOODLIGHT_TAGS, 'I1').value == 'MetroPCS':
            adv_id = self.configs['advertiser_ids']['MetroPCS']
        elif Range(DCM_API().GENERATE_FLOODLIGHT_TAGS, 'I1').value == 'Both':
            adv_id = [self.configs['advertiser_ids']['T-Mobile'],
                      self.configs['advertiser_ids']['MetroPCS']]
        else:
            sys.exit('An Advertiser must be selected')

        floodlight_tags = pd.DataFrame()

        if Range(self.GENERATE_FLOODLIGHT_TAGS, 'I1').value == 'T-Mobile' or 'MetroPCS':
            floodlights = self.fl
            request = floodlights.list(profileId=self.prof_id,
                                       advertiserId=adv_id,
                                       fields='floodlightActivities(id,expectedUrl,floodlightActivityGroupId,' +
                                              'floodlightActivityGroupName,name,tagString)')

            while request is not None:
                tags = request.execute()

                try:
                    tags_norm = json_normalize(tags['floodlightActivities'])
                except IndexError:
                    break

                floodlight_tags = floodlight_tags.append(tags_norm)
                request = floodlights.list_next(request, tags)

        else:
            metro_floodlights = self.service.floodlightActivities().list(profileId=self.prof_id,
                                                                         advertiserId=adv_id[1]).execute()
            tmo_floodlights = self.service.floodlightActivities().list(profileId=self.prof_id,
                                                                       advertiserId=adv_id[0]).execute()

            metro_floodlights = json_normalize(metro_floodlights['floodlightActivities'])
            tmo_floodlights = json_normalize(tmo_floodlights['floodlightActivities'])
            floodlights_json = tmo_floodlights.append(metro_floodlights)

        floodlight_tags = floodlight_tags[['id', 'name', 'tagString', 'expectedUrl',
                                           'floodlightActivityGroupId', 'floodlightActivityGroupName']]
        floodlight_tags.drop_duplicates(inplace=True)

        generated_tags = []
        batch_tag_generation = self.service.new_batch_http_request()

        def list_tags(request_id, response, exception):
            if exception is not None:
                pass
            else:
                tag = response
                try:
                    tag = tag['floodlightActivityTag']
                except TypeError:
                    tag = 'Could Not Load'

                generated_tags.append(tag)

        for index, row in floodlight_tags.iterrows():
            batch_tag_generation.add(self.fl.generatetag(profileId=self.prof_id,
                                                                                floodlightActivityId=str(row[0]),
                                                                                fields='floodlightActivityTag'),
                                     callback=list_tags)

        http = httplib2.Http()
        batch_tag_generation.execute(http=http)

        floodlight_tags['Tag'] = generated_tags

        DataMethods().chunk_df(floodlight_tags, self.GENERATE_FLOODLIGHT_TAGS, 'I3')

    def generate_selected_tags(self):
        tags = pd.DataFrame(Range(self.GENERATE_FLOODLIGHT_TAGS, 'A3').table.value,
                            columns=Range(self.GENERATE_FLOODLIGHT_TAGS, 'A3').horizontal.value)

        tags.drop(0, inplace=True)

        tag_name = []
        floodlight_id = []
        generated_tag = []

        for index, row in tags.iterrows():
            tag = self.fl.generatetag(profileId=self.prof_id,
                                                             floodlightActivityId=str(row[0])).execute()

            floodlight_id.append(row[0])
            tag_name.append(row[1])
            generated_tag.append(tag['floodlightActivityTag'])

        tag_info = [floodlight_id, tag_name, generated_tag]
        tag_info = {i[0]: list(i[1:]) for i in zip(*tag_info)}

        tag_df = pd.DataFrame.from_dict(tag_info, orient='index').reset_index()
        tag_df.rename(columns={'index': 'Floodlight ID', 0: 'Floodlight Name', 1: 'Tag'}, inplace=True)

        DataMethods().chunk_df(tag_df, self.GENERATE_FLOODLIGHT_TAGS, 'H1')
