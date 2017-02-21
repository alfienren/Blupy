import pandas as pd
from xlwings import Range

from analytics.data_refresh.data import DataMethods
from dcm.dcm_api import DCM_API


class Pixels(DCM_API):

    def __init__(self):
        super(Pixels, self).__init__()

    def get(self):
        floodlights = Range(self.LIST_FLOODLIGHT_PIXELS, 'A2').vertical.value

        df = pd.DataFrame()
        for i in floodlights:
            try:
                pix = self.fl.get(profileId=self.prof_id, id=str(int(i))).execute()[
                    'defaultTags']
            except:
                pass

            tags = pd.DataFrame(pix)
            del tags['id']
            tags['Floodlight ID'] = i
            df = df.append(tags)

        df = df[['Floodlight ID', 'name', 'tag']]

        DataMethods().chunk_df(df, self.LIST_FLOODLIGHT_PIXELS, 'K1')

    def implement(self):
        pixels = pd.DataFrame(Range(self.PIGGYBACK_PIXELS, 'A1').table.value,
                              columns=Range(self.PIGGYBACK_PIXELS, 'A1').horizontal.value)

        pixels.drop(0, inplace=True)

        ids = list(pixels['Floodlight ID'].unique())

        grouped = pixels.groupby(pixels['Floodlight ID'])

        for i in ids:
            p = self.fl.get(profileId=self.prof_id, id=str(i)).execute()
            for index, row in grouped.get_group(i).iterrows():
                try:
                    p['defaultTags'].append({'name': row['Pixel Name'], 'tag': row['Pixel Tag']})
                except KeyError:
                    p['defaultTags'] = [

                    ]
                    p['defaultTags'].append({'name': row['Pixel Name'], 'tag': row['Pixel Tag']})

            req = {
                'defaultTags':
                    p['defaultTags']
            }

            self.fl.patch(profileId=self.prof_id, id=i, body=req).execute()

            # try:
            #
            # except HttpError:
            #     #service.floodlightActivities().patch(profileId=prof_id, id=i, body=req_null).execute()
            #     req_null = req['defaultTags']
            #     for j in req_null:
            #         if 'id' in j:
            #             del j['id']
            #         j['name'] = None
            #         j['tag'] = None
            #
            #     service.floodlightActivities().patch(profileId=prof_id, id=i, body=req_null).execute()
            #     service.floodlightActivities().patch(profileId=prof_id, id=i, body=req).execute()
