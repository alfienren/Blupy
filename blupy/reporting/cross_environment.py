import numpy as np
import pandas as pd
from data.file_io import DataMethods
from xlwings import Range, Sheet

from floodlights import Floodlights


class CrossEnvironment(DataMethods):

    def __init__(self):
        super(CrossEnvironment, self).__init__()

    def cross_environment(self):
        piv = pd.read_excel(Range(self.action_reference, 'AE1').value, 'data', parse_cols='B:AY', index_col=None)
        pivoted_data = pd.pivot_table(piv,
                                      index=['Date', 'Campaign', 'Media Plan', 'Site (DCM)', 'Site', 'Tactic',
                                             'Desktop/Mobile/Video', 'Sub - Tactic', 'Placement', 'Placement ID'],
                                      aggfunc=np.sum).reset_index()
        pivoted_data.set_index(['Date', 'Site (DCM)', 'Placement', 'Placement ID'], inplace=True)

        data = pd.read_excel(self.path, 'raw', index_col=None)
        data = Floodlights().a_e_traffic(data, adv='tmo')

        a_actions, b_actions, c_actions, d_actions, e_actions = Range('Action_Reference', 'A2').vertical.value, Range(
            'Action_Reference', 'B2').vertical.value, Range('Action_Reference', 'C2').vertical.value, Range(
            'Action_Reference', 'D2').vertical.value, Range('Action_Reference', 'E2').vertical.value

        column_names = data.columns

        a_actions_ce, b_actions_ce, c_actions_ce, d_actions_ce, e_actions_ce = [], [], [], [], []

        for i, j in zip([a_actions_ce, b_actions_ce, c_actions_ce, d_actions_ce, e_actions_ce],
                        [a_actions, b_actions, c_actions, d_actions, e_actions]):
            i.extend([s + ' + Cross-Environment' for s in j])

        a_actions_ce, b_actions_ce, c_actions_ce, d_actions_ce, e_actions_ce = list(
            set(a_actions_ce).intersection(column_names)), \
                                                                               list(set(b_actions_ce).intersection(
                                                                                   column_names)), \
                                                                               list(set(c_actions_ce).intersection(
                                                                                   column_names)), \
                                                                               list(set(d_actions_ce).intersection(
                                                                                   column_names)), \
                                                                               list(set(e_actions_ce).intersection(
                                                                                   column_names))

        data['A Actions + Cross-Environment'], data['B Actions + Cross-Environment'], data[
            'C Actions + Cross-Environment'], \
        data['D Actions + Cross-Environment'], data['E Actions + Cross-Environment'] = \
            data[a_actions_ce].sum(axis=1), data[b_actions_ce].sum(axis=1), data[c_actions_ce].sum(axis=1), \
            data[d_actions_ce].sum(axis=1), data[e_actions_ce].sum(axis=1)

        data['Standard Conversions Traffic Actions'] = data['A Actions'] + data['B Actions'] + data['C Actions'] + \
                                                       data['D Actions']

        data['Standard + Cross Conversions Traffic Actions'] = data['A Actions + Cross-Environment'] + data[
            'B Actions + Cross-Environment'] + data['C Actions + Cross-Environment'] + data[
                                                                   'D Actions + Cross-Environment'] + data[
                                                                   'E Actions + Cross-Environment']

        data_pivoted = pd.pivot_table(data,
                                      index=['Date', 'Campaign', 'Site (DCM)', 'Placement', 'Placement ID',
                                             'Platform Type'],
                                      aggfunc=np.sum).reset_index()
        data_pivoted.set_index(['Date', 'Site (DCM)', 'Placement', 'Placement ID'], inplace=True)
        merged = pd.merge(data_pivoted, pivoted_data, right_index=True, left_index=True, how='left').reset_index()
        merged['key'] = merged['Date'].astype(str) + merged['Placement']
        unique_keys = list(merged['key'].unique())

        for i in unique_keys:
            merged['NTC Media Cost'] = np.where(merged['key'] == i,
                                                merged['NTC Media Cost'] / len(merged[merged['key'] == i]),
                                                merged['NTC Media Cost'])

            merged['Impressions'] = np.where(merged['key'] == i,
                                             merged['Impressions'] / len(merged[merged['key'] == i]),
                                             merged['Impressions'])

            merged['Clicks'] = np.where(merged['key'] == i,
                                        merged['Clicks'] / len(merged[merged['key'] == i]),
                                        merged['Clicks'])

        dim_columns = ['Date', 'Campaign', 'Media Plan', 'Site', 'Site (DCM)', 'Tactic', 'Sub - Tactic',
                       'Desktop/Mobile/Video', 'Placement', 'Placement ID']

        metric_columns = ['NTC Media Cost', 'Media Cost', 'Orders', 'eGAs', 'Impressions', 'Clicks', 'A Actions',
                          'B Actions', 'C Actions', 'D Actions', 'A Actions + Cross-Environment',
                          'B Actions + Cross-Environment', 'C Actions + Cross-Environment',
                          'D Actions + Cross-Environment', 'E Actions', 'E Actions + Cross-Environment']

        merged.rename(columns={'Campaign_x': 'Campaign'}, inplace=True)
        keep_columns = dim_columns + metric_columns

        for col in list(merged.columns):
            if col not in keep_columns:
                merged.drop(col, axis=1, inplace=True)

        merged = merged[keep_columns]

        Sheet('data').clear_contents()
        DataMethods().chunk_df(merged, 'data', 'A1')