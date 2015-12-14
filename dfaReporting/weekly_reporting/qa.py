import numpy as np
from xlwings import Sheet, Range

def placement_qa(data):

    # Check placements with spend and impressions but 0 clicks. Also check for placements with spend and clicks but
    # no impressions.

    data['Flag'] = np.where((data['Media Cost'] > 10) &
                                 (data['Impressions'] > 100) &
                                 (data['Clicks'] == 0), 'Zero Clicks',
                                 np.where((data['Media Cost'] > 0) &
                                          (data['Clicks'] > 0) &
                                          (data['Impressions'] == 0), 'Zero Impressions', np.nan))

    data = data[data['Flag'] != 'nan']

    data = data[['Placement', 'Placement ID', 'Date', 'Media Cost', 'Impressions', 'Clicks', 'Flag']]

    Sheet.add('Data_QA_Output', after = 'data')

    Range('Data_QA_Output', 'A1', index = False).value = data


