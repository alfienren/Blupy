import pandas as pd
import numpy as np
from xlwings import Sheet, Range

def placement_qa(data):

    # Check placements with spend and impressions but 0 clicks. Also check for placements with spend and clicks but
    # no impressions.

    data_qa = data[['Placement', 'Placement ID', 'Date', 'Media Cost', 'Impressions', 'Clicks']]

    data_qa['Flag'] = np.where((data_qa['Media Cost'] > 10) &
                                 (data_qa['Impressions'] > 100) &
                                 (data_qa['Clicks'] == 0), 'Zero Clicks',
                                 np.where((data_qa['Media Cost'] > 0) &
                                          (data_qa['Clicks'] > 0) &
                                          (data_qa['Impressions'] == 0), 'Zero Impressions', np.nan))

    data_qa = data_qa[data_qa['Zeroes'] != 'nan']

    Sheet.add('Data_QA_Output', after = 'data')

    Range('Data_QA_Output').value = data_qa


