import pandas as pd
import numpy as np


def merge_channel_data(online, offline, search, social, tmo):
    online_combined = online.append([social, search])

    online_combined['Sub-Tactic'] = np.where(online_combined['Sub-Tactic'].isnull() == True, online_combined['Medium'],
                                             online_combined['Sub-Tactic'])
    online_combined['Tactic'] = np.where(online_combined['Tactic'].isnull() == True, online_combined['Medium'],
                                         online_combined['Tactic'])

    tmo_inputs = pd.pivot_table(tmo, index=['Week'], columns=['Metric'], values=['Volume'], aggfunc=np.sum)
    tmo_inputs.columns = tmo_inputs.columns.get_level_values(1)

    search_impressions = pd.pivot_table(search, index=['Week'], values=['Branded Search Impressions'],
                                        aggfunc=np.sum).reset_index()

    online_offline = online_combined.append(offline)
    online_offline.set_index('Week', inplace=True)

    online_offline_tmo = pd.merge(online_offline, tmo_inputs, how='left',
                                  right_index=True, left_index=True).reset_index()

    weeks = list(online_offline_tmo['Week'].unique())
    metric_columns = ['Customer Traffic', 'Direct Load Traffic', 'Gross Adds', 'Mobile Visits',
                      'Non-Mobile Visits', 'Prospect Traffic', 'Retail Traffic', 'Total Orders', 'Total Traffic']

    for i in weeks:
        for j in metric_columns:
            online_offline_tmo[j] = np.where(online_offline_tmo['Week'] == i,
                                             online_offline_tmo[j] / len(online_offline_tmo[online_offline_tmo['Week'] == i]),
                                             online_offline_tmo[j])

    aggregated_tmo = pd.pivot_table(tmo, index=['Metric', 'Week'], aggfunc=np.sum).reset_index()
    search_impressions['Metric'] = 'Branded Search'
    search_impressions.rename(columns={'Branded Search Impressions':'Volume'}, inplace=True)

    aggregated_tmo = aggregated_tmo.append(search_impressions)
    online_offline.reset_index(inplace=True)

    aggregated_tmo.set_index('Week', inplace=True)
    spend = pd.pivot_table(online_offline, index=['Week'], values=['Spend'], aggfunc=np.sum)
    aggregated_tmo = pd.merge(aggregated_tmo, spend, how='left', right_index=True, left_index=True).reset_index()

    weeks = list(aggregated_tmo['Week'].unique())

    for i in weeks:
        aggregated_tmo['Spend'] = np.where(aggregated_tmo['Week'] == i,
                                           aggregated_tmo['Spend'] / len(aggregated_tmo[aggregated_tmo['Week'] == i]),
                                           aggregated_tmo['Spend'])

    return [online_offline_tmo, aggregated_tmo]