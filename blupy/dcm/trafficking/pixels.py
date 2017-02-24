import pandas as pd
from xlwings import Range
import httplib2

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

    def delete(self):
        pixels = pd.DataFrame(Range(self.DELETE_PIXELS, 'A1').table.value,
                              columns=Range(self.DELETE_PIXELS, 'A1').horizontal.value)
        pixels.drop(0, inplace=True)

        ids = list(pixels['Floodlight ID'].unique())

        grouped = pixels.groupby(pixels['Floodlight ID'])
        delete_pixels = self.service.new_batch_http_request()
        add_pixels = self.service.new_batch_http_request()

        empty_pixel_patch = {
            "defaultTags": [

            ]
        }

        http = httplib2.Http()

        for i in ids:
            p = self.fl.get(profileId=self.prof_id, id=str(i)).execute()
            delete_pixels.add(self.fl.patch(profileId=self.prof_id, id=str(i), body=empty_pixel_patch))

            try:
                pix = p['defaultTags']
            except KeyError:
                pix = []

            if pix is not []:
                for index, row in grouped.get_group(i).iterrows():
                    for j in range(0, len(pix)):
                        try:
                            if pix[j]['name'] == row['Pixel Name']:
                                pix.pop(j)
                        except IndexError:
                            pass

                for k in pix:
                    k.pop('id')

            new_pixel_body = {
                "defaultTags": pix
            }

            add_pixels.add(self.fl.patch(profileId=self.prof_id, id=str(i), body=new_pixel_body))

        delete_pixels.execute(http=http)
        add_pixels.execute(http=http)
