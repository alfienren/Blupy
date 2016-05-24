import numpy as np
from xlwings import Sheet, Range
from reporting.constants import TabNames

# Check placements with spend and impressions but 0 clicks. Also check for placements with spend and clicks but
# no impressions.

def placement_qa(data):
    data['Flag'] = np.where((data['Media Cost'] > 50) & (data['Impressions'] < 1000), 'Low Impressions',
                            np.where((data['Media Cost'] > 10) & (data['Impressions'] > 100) & (data['Clicks'] == 0),
                                     'Zero Clicks',
                                     np.where(
                                         (data['Media Cost'] > 0) & (data['Clicks'] > 0) & (data['Impressions'] == 0),
                                         'Zero Impressions', np.nan)))

    data = data[data['Flag'] != 'nan']
    data = data[['Placement', 'Placement ID', 'Week', 'Media Cost', 'Impressions', 'Clicks', 'Flag']]

    Sheet.add('Data_QA_Output', after = 'data')

    Range(TabNames.qa_tab_name, 'A1', index = False).value = data


