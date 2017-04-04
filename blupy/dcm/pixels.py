import httplib2
import pandas as pd
from xlwings import Range

from dcm.dcm_api import DCM_API
from data.file_io import DataMethods


class Pixels(DCM_API):

    def __init__(self):
        super(Pixels, self).__init__()

    def get(self):
        floodlights = Range(self.LIST_FLOODLIGHT_PIXELS, 'A2').vertical.value
        batch_floodlights = self.service.new_batch_http_request()

        tags = []

        def floodlight_callback(request_id, response, exception):
            if exception is not None:
                tag = pd.DataFrame({'Floodlight ID': {0: floodlights[int(request_id)-1]},
                                    'name': {0: 'None'},
                                    'tag': {0: 'None'}})
                tags.append(tag)
            else:
                try:
                    r = response['defaultTags']
                    tag = pd.DataFrame(r)
                    del tag['id']
                    tag['Floodlight ID'] = floodlights[int(request_id)-1]
                except:
                    tag = pd.DataFrame({'Floodlight ID': {0: floodlights[int(request_id)-1]},
                                        'name': {0: 'None'},
                                        'tag': {0: 'None'}})

                tags.append(tag)

        for i in floodlights:
            batch_floodlights.add(self.fl.get(profileId=self.prof_id, id=i), callback=floodlight_callback)

        batch_floodlights.execute()

        df = pd.DataFrame()
        for j in tags:
            df = df.append(j)

        df = df[['Floodlight ID', 'name', 'tag']]
        DataMethods().chunk_df(df, self.LIST_FLOODLIGHT_PIXELS, 'K1')

    def implement(self):
        pixels = pd.DataFrame(Range(self.PIGGYBACK_PIXELS, 'A1').table.value,
                              columns=Range(self.PIGGYBACK_PIXELS, 'A1').horizontal.value)
        pixels.drop(0, inplace=True)
        batch_pixels = self.service.new_batch_http_request()

        def implement_pixel_callback(request_id, response, exception):
            if exception is not None:
                pass
            else:
                pass

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
