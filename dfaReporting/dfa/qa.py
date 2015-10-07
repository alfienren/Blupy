import pandas as pd
import numpy as np

def placement_qa(data):

    # Check placements with spend and impressions but 0 clicks. Also check for placements with spend and clicks but
    # no impressions.

    data_qa = data[['Placement', 'Placement ID', 'Date', 'Media Cost', 'Impressions', 'Clicks']]

    data_qa['Zero Clicks'] = np.where((data_qa['Media Cost'] > 0) &
                                      (data_qa['Impressions'] > 0) &
                                      (data_qa['Clicks'] == 0), data_qa['Placement'], np.nan)

    data_qa['Zero Impressions'] = np.where((data_qa['Media Cost'] > 0) &
                                           (data_qa['Clicks'] > 0) &
                                           (data_qa['Impressions'] == 0), data_qa['Placement'], np.nan)

    zero_clicks = data_qa[['Placement', ]]

