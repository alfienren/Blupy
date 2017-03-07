import re

import numpy as np
import pandas as pd
from xlwings import Range


class Floodlights(object):

    def __init__(self):
        self.view_through = .25
        self.action_reference = 'Action_Reference'
        self.spanish_language = '|'.join(list(['Spanish', 'Hispanic', 'SL']))
        self.device_feed = pd.read_table(Range(self.action_reference, 'AE1').value)

    def a_e_traffic(self, data, adv='tmo'):
        # Set the data column names to a variable
        column_names = data.columns

        a, b, c, d, e = [], [], [], [], []
        funnel_tags = [a, b, c, d, e]

        col = chr(ord('A'))

        for i in range(0, len(funnel_tags)):
            ref = Range(self.action_reference, col + '1').vertical.value
            funnel_tags[i].append(ref)
            col = chr(ord(str(col)) + 1)

        # For each action tag category (A - E), search the column names to find the action tag. A new list of the A - E
        # actions are compiled from the matches in the column names.
        a_actions, b_actions, c_actions, d_actions, e_actions = list(set(a[0L]).intersection(column_names)), \
                                                                list(set(b[0L]).intersection(column_names)), \
                                                                list(set(c[0L]).intersection(column_names)), \
                                                                list(set(d[0L]).intersection(column_names)), \
                                                                list(set(e[0L]).intersection(column_names))

        # With the references set for each action tag category, sum the tags along rows and create columns for each action
        # tag bucket.

        data['A Actions'], data['B Actions'], data['C Actions'], data['D Actions'], data['E Actions'] = data[
                                                                                                            a_actions].sum(
            axis=1), \
                                                                                                        data[
                                                                                                            b_actions].sum(
                                                                                                            axis=1), \
                                                                                                        data[
                                                                                                            c_actions].sum(
                                                                                                            axis=1), \
                                                                                                        data[
                                                                                                            d_actions].sum(
                                                                                                            axis=1), \
                                                                                                        data[
                                                                                                            e_actions].sum(
                                                                                                            axis=1)

        # Similar to the logic to get the sum of action tags for each respective category, the following finds the sum
        # of post-click and impression activity as well as Store Locator Visits.

        # For post-view and click activity, the column names of the data are searched for columns containing 'View-through',
        # 'Click-through' or 'Store Locator'. Column matches are then stored in lists.

        view_through, click_through, store_locator = [], [], []

        for i in [view_through, click_through, store_locator]:
            for item in column_names:
                view = re.search('View-through Conversions', item)
                click = re.search('Click-through Conversions', item)
                locator = re.search('Store Locator 2', item)
                if view:
                    i.append(item)
                if click:
                    i.append(item)
                if locator:
                    i.append(item)

        # The matches against the column names are then set
        view_based, click_based, slv_conversions = list(set(view_through).intersection(column_names)), \
                                                   list(set(click_through).intersection(column_names)), \
                                                   list(set(store_locator).intersection(column_names))

        # With the matching references set, sum the matches for post-click and impression activity as well as SLV by row,
        # creating columns for each.
        data['Post-Click Activity'], data['Post-Impression Activity'], data['Store Locator Visits'] = data[
                                                                                                          click_based].sum(
            axis=1), data[view_based].sum(axis=1), data[slv_conversions].sum(axis=1)

        if adv != 'tmo':
            lat_a, lat_b, lat_c, lat_d, lat_e, gm_a, gm_b, gm_c, gm_d, gm_e = \
            [], [], [], [], [], [], [], [], [], []

            for i in a_actions:
                lat = re.search(self.spanish_language, i)
                if lat:
                    lat_a.append(i)
                if not lat:
                    gm_a.append(i)

            for i in b_actions:
                lat = re.search(self.spanish_language, i)
                if lat:
                    lat_b.append(i)
                if not lat:
                    gm_b.append(i)

            for i in c_actions:
                lat = re.search(self.spanish_language, i)
                if lat:
                    lat_c.append(i)
                if not lat:
                    gm_c.append(i)

            for i in d_actions:
                lat = re.search(self.spanish_language, i)
                if lat:
                    lat_d.append(i)
                if not lat:
                    gm_d.append(i)

            for i in e_actions:
                lat = re.search(self.spanish_language, i)
                if lat:
                    lat_e.append(i)
                if not lat:
                    gm_e.append(i)

            gm_a, gm_b, gm_c, gm_d, gm_e = list(set(gm_a).intersection(column_names)), \
                                           list(set(gm_b).intersection(column_names)), \
                                           list(set(gm_c).intersection(column_names)), \
                                           list(set(gm_d).intersection(column_names)), \
                                           list(set(gm_e).intersection(column_names))

            lat_a, lat_b, lat_c, lat_d, lat_e = list(set(lat_a).intersection(column_names)), \
                                                list(set(lat_b).intersection(column_names)), \
                                                list(set(lat_c).intersection(column_names)), \
                                                list(set(lat_d).intersection(column_names)), \
                                                list(set(lat_e).intersection(column_names))

            data['GM A Actions'], data['GM B Actions'], data['GM C Actions'], data['GM D Actions'], data['GM E Actions'] = \
            data[gm_a].sum(axis=1), \
            data[gm_b].sum(axis=1), \
            data[gm_c].sum(axis=1), \
            data[gm_d].sum(axis=1), \
            data[gm_e].sum(axis=1)

            data['Hispanic A Actions'], data['Hispanic B Actions'], data['Hispanic C Actions'], data['Hispanic D Actions'], \
            data['Hispanic E Actions'] = data[lat_a].sum(axis=1), \
                                         data[lat_b].sum(axis=1), \
                                         data[lat_c].sum(axis=1), \
                                         data[lat_d].sum(axis=1), \
                                         data[lat_e].sum(axis=1)

            data['Total A Actions'], data['Total B Actions'], data['Total C Actions'], data['Total D Actions'], data[
                'Total E Actions'] = data['GM A Actions'] + data['Hispanic A Actions'], \
                                     data['GM B Actions'] + data['Hispanic B Actions'], \
                                     data['GM C Actions'] + data['Hispanic C Actions'], \
                                     data['GM D Actions'] + data['Hispanic D Actions'], \
                                     data['GM E Actions'] + data['Hispanic E Actions']

            data['Orders'] = data['Transaction Count']

        # Create columns for Awareness and Consideration actions. Awareness Actions are the sum of A and B actions,
        # Consideration is the sum of C and D actions
        # Traffic Actions are the total of Awareness and Consideration Actions, or A - D actions.
        data['Awareness Actions'], data['Consideration Actions'] = data['A Actions'] + data['B Actions'], \
                                                                   data['C Actions'] + data['D Actions']
        data['Traffic Actions'] = data['Awareness Actions'] + data['Consideration Actions']

        return data

    def custom_variables(self, cfv):
        device_feed = self.device_feed

        prepaid, postpaid = device_feed[device_feed['Product Subcategory'] == 'Prepaid'], \
                            device_feed[device_feed['Product Subcategory'] == 'Postpaid']

        prepaid_list, postpaid_list = prepaid['Device SKU'].tolist(), postpaid['Device SKU'].tolist()
        prepaid_list, postpaid_list = '|'.join(prepaid_list), '|'.join(postpaid_list)

        cfv['Device_reg'] = cfv['Device (string)'].apply(lambda x: str(x).replace(',', '|'))

        # Count the number of plans in the Plans column
        cfv['Plans'] = np.where(cfv['Plan (string)'] != np.NaN,
                                cfv['Plan (string)'].apply(lambda x: str(x).count(',')) + 1, 0)
        # Count number of services in the Service column
        cfv['Services'] = np.where(cfv['Service (string)'] != np.NaN,
                                   cfv['Service (string)'].apply(lambda x: str(x).count(',')) + 1, 0)
        # Count number of Accessories in the Accessories column
        cfv['Accessories'] = np.where(cfv['Accessory (string)'] != np.NaN,
                                      cfv['Accessory (string)'].apply(lambda x: str(x).count(',')) + 1, 0)
        # Count number of devices in the Plans column
        cfv['Devices'] = np.where(cfv['Device (string)'] != np.NaN,
                                  cfv['Device (string)'].apply(lambda x: str(x).count(',')) + 1, 0)
        # Count number of Add-a-Lines in the Service column
        cfv['Add-a-Line'] = cfv['Service (string)'].apply(lambda x: str(x).count('ADD'))
        # Activations are defined as the sum of Plans and Add-a-Line
        cfv['Activations'] = cfv['Plans'] + cfv['Add-a-Line']

        # Postpaid plans are defined as a Plan + Device. By row, if number of plans is equal to number of devices, Postpaid
        # plans = number of plans. If plans and devices are not equal, use the minimum number.
        cfv['Postpaid Plans'] = abs(np.where(cfv['Plans'].fillna(0) == cfv['Devices'].fillna(0), cfv['Plans'],
                                             np.where((cfv['Plans'].fillna(0) > cfv['Devices'].fillna(0)) | (
                                                 cfv['Plans'].fillna(0) < cfv['Devices'].fillna(0)),
                                                      pd.concat([cfv['Plans'].fillna(0), cfv['Devices'].fillna(0)],
                                                                axis=1).min(axis=1), 0)))

        # Prepaid plans are defined as the number of Devices with no service and plan. If number of plans and services are
        # 0, count of devices is prepaid. If service and plan are not equal, prepaid plans = 0.
        cfv['Prepaid Plans'] = abs(
            np.where((cfv['Plans'].fillna(0) == 0) & (cfv['Services'].fillna(0) == 0), cfv['Devices'],
                     np.where((cfv['Devices'].fillna(0) > cfv['Plans'].fillna(0)) & (
                         cfv['Devices'].fillna(0) > cfv['Services'].fillna(0)),
                              cfv['Devices'].fillna(0) - pd.concat([cfv['Plans'].fillna(0), cfv['Services'].fillna(0)],
                                                                   axis=1).max(
                                  axis=1), 0)))

        # The DDR campaign counts view-through order credit at 50%. If the campaign name contains 'DDR' and the Floodlight
        # Attribution Type is View-through, the order is multiplied by 0.5.
        # cfv['Orders'] = np.where(((cfv['Plan (string)'].isnull() == True) & (cfv['Service (string)'].isnull() == True) & (
        # cfv['Device (string)'].isnull() == True) & (cfv['Accessory (string)'].notnull() == True)), 0,
        #                          np.where(cfv['Floodlight Attribution Type'].str.contains('View-through') == True,
        #                                   cfv['Orders'] * view_through_credit(), cfv['Orders']))

        cfv['Postpaid Orders'] = np.where((cfv['Device_reg'].str.contains(postpaid_list) == True) & (
            cfv['Floodlight Attribution Type'].str.contains('View-through') == True),
                                          1 * self.view_through, np.where(
                (cfv['Device_reg'].str.contains(postpaid_list) == True) & (
                    cfv['Floodlight Attribution Type'].str.contains('Click-through') == True),
                1, 0))

        cfv['Prepaid Orders'] = np.where((cfv['Device_reg'].str.contains(prepaid_list) == True) & (
            cfv['Floodlight Attribution Type'].str.contains('View-through') == True),
                                         1 * self.view_through, np.where(
                (cfv['Device_reg'].str.contains(prepaid_list) == True) & (
                    cfv['Floodlight Attribution Type'].str.contains('Click-through') == True),
                1, 0))

        cfv['Orders'] = cfv['Postpaid Orders'] + cfv['Prepaid Orders']

        # Estimated Gross Adds are calculated as the count of Devices with 50% view-through credit.
        # If Floodlight Attribution Type is equal to View-through, the count of Devices is multiplied by 0.5
        cfv['eGAs'] = np.where(cfv['Floodlight Attribution Type'].str.contains('View-through') == True,
                               (cfv['Device (string)'].apply(lambda x: str(x).count(',')) + 1) * self.view_through,
                               (cfv['Device (string)'].apply(lambda x: str(x).count(',')) + 1))

        return cfv

    def ddr_custom_variables(self, cfv):
        device_string = cfv['Device (string)'].apply(lambda x: str(x).split(',')).apply(pd.Series).stack()
        device_string.index = device_string.index.droplevel(-1)
        device_string.name = "Device IDs"

        device_cfv = cfv[cfv.columns[0:17]].join(device_string)
        cfv = cfv.append(device_cfv)

        excluded_devices = str(Range('Lookup', 'S2').value)
        cfv = pd.merge(cfv, self.device_feed, how='left', left_on='Device IDs', right_on='Device SKU')

        cfv['Prepaid GAs'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                       (cfv['Device IDs'].notnull() == True) & (
                                           (cfv['Product Subcategory'].str.contains('Prepaid') == True) | (
                                               cfv['Device IDs'].notnull()) & (cfv['Product Subcategory'].isnull())) &
                                       (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)),
                                      self.view_through,
                                      np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) & (
                                          cfv['Device IDs'].notnull() == True) &
                                                (cfv['Product Subcategory'].str.contains('Prepaid') == True) &
                                                (cfv['Floodlight Attribution Type'].str.contains(
                                                    'Click-through') == True)),
                                               1, 0))

        cfv['Postpaid GAs'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                        (cfv['Device IDs'].notnull() == True) & (
                                            cfv['Product Subcategory'].str.contains('Postpaid') == True) &
                                        (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)),
                                       self.view_through,
                                       np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                                 (cfv['Device IDs'].notnull() == True) & (
                                                     cfv['Product Subcategory'].str.contains('Postpaid') == True)), 1,
                                                0))

        cfv['Total GAs'] = cfv['Postpaid GAs'] + cfv['Prepaid GAs']

        cfv['Prepaid SIMs'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                        (cfv['Device IDs'].notnull() == True) & (
                                            cfv['Product Category'].str.contains('SIM card') == True) &
                                        (cfv['Product Subcategory'].str.contains('Prepaid') == True) &
                                        (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)),
                                       self.view_through,
                                       np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                                 (cfv['Device IDs'].notnull() == True) & (
                                                     cfv['Product Category'].str.contains('SIM card') == True) &
                                                 (cfv['Product Subcategory'].str.contains('Prepaid') == True) &
                                                 (
                                                     cfv['Floodlight Attribution Type'].str.contains(
                                                         'Click-through') == True)),
                                                1, 0))

        cfv['Postpaid SIMs'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                         (cfv['Device IDs'].notnull() == True) & (
                                             cfv['Product Category'].str.contains('SIM card') == True) &
                                         (cfv['Floodlight Attribution Type'].str.contains('View-through') == True) &
                                         (cfv['Product Subcategory'].str.contains('Postpaid') == True)),
                                        self.view_through,
                                        np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                                  (cfv['Device IDs'].notnull() == True) & (
                                                      cfv['Product Category'].str.contains('SIM card') == True) &
                                                  (cfv['Product Subcategory'].str.contains('Postpaid') == True) &
                                                  (cfv['Floodlight Attribution Type'].str.contains(
                                                      'Click-through') == True)), 1, 0))

        cfv['Prepaid Mobile Internet'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                                   (cfv['Device IDs'].notnull() == True) & (
                                                       cfv['Product Category'].str.contains(
                                                           'Mobile Internet') == True) &
                                                   (cfv['Product Subcategory'].str.contains('Prepaid') == True) &
                                                   (cfv['Floodlight Attribution Type'].str.contains(
                                                       'View-through') == True)), self.view_through,
                                                  np.where(
                                                      ((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                                       (cfv['Device IDs'].notnull() == True) & (
                                                           cfv['Product Category'].str.contains(
                                                               'Mobile Internet') == True) &
                                                       (cfv['Product Subcategory'].str.contains('Prepaid') == True)),
                                                      1, 0))

        cfv['Postpaid Mobile Internet'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                                    (cfv['Device IDs'].notnull() == True) & (
                                                        cfv['Product Category'].str.contains(
                                                            'Mobile Internet') == True) &
                                                    (cfv['Product Subcategory'].str.contains('Postpaid') == True) &
                                                    (cfv['Floodlight Attribution Type'].str.contains(
                                                        'View-through') == True)), self.view_through,
                                                   np.where(
                                                       ((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                                        (cfv['Device IDs'].notnull() == True) & (
                                                            cfv['Product Category'].str.contains(
                                                                'Mobile Internet') == True) &
                                                        (cfv['Product Subcategory'].str.contains('Postpaid') == True) &
                                                        (cfv['Floodlight Attribution Type'].str.contains(
                                                            'Click-through') == True)), 1, 0))

        cfv['Prepaid Phone'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                         (cfv['Device IDs'].notnull() == True) & (
                                             cfv['Product Category'].str.contains('Smartphone') == True) &
                                         (cfv['Product Subcategory'].str.contains('Prepaid') == True) &
                                         (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)),
                                        self.view_through,
                                        np.where((((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                                   (cfv['Device IDs'].notnull() == True) & (
                                                       cfv['Product Category'].str.contains('Smartphone') == True) &
                                                   (cfv['Product Subcategory'].str.contains('Prepaid') == True) &
                                                   (cfv['Floodlight Attribution Type'].str.contains(
                                                       'Click-through') == True))), 1, 0))

        cfv['Postpaid Phone'] = np.where(((cfv['Device IDs'].str.contains(excluded_devices) == False) &
                                          (cfv['Device IDs'].notnull() == True) & (
                                              cfv['Product Category'].str.contains('Smartphone') == True) &
                                          (cfv['Product Subcategory'].str.contains('Postpaid') == True) &
                                          (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)),
                                         self.view_through,
                                         np.where(((cfv['Device IDs'].notnull() == True) & (
                                             cfv['Product Category'].str.contains('Smartphone') == True) &
                                                   (cfv['Product Subcategory'].str.contains('Postpaid') == True) &
                                                   (cfv['Floodlight Attribution Type'].str.contains(
                                                       'Click-through') == True)), 1, 0))

        cfv['DDR New Devices'] = np.where(
            ((cfv['Device IDs'].str.contains(excluded_devices) == False) & (cfv['Device IDs'].notnull() == True) &
             (cfv['Activity'].str.contains('New TMO Order') == True) &
             (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)), self.view_through,
            np.where(
                ((cfv['Device IDs'].str.contains(excluded_devices) == False) & (cfv['Device IDs'].notnull() == True) &
                 (cfv['Activity'].str.contains('New TMO Order') == True) &
                 (cfv['Floodlight Attribution Type'].str.contains('Click-through') == True)), 1, 0))

        cfv['DDR Add-a-Line'] = np.where(
            ((cfv['Device IDs'].str.contains(excluded_devices) == False) & (cfv['Device IDs'].notnull() == True) &
             (cfv['Activity'].str.contains('New My.TMO Order') == True) &
             (cfv['Floodlight Attribution Type'].str.contains('View-through') == True)), self.view_through,
            np.where(
                ((cfv['Device IDs'].str.contains(excluded_devices) == False) & (cfv['Device IDs'].notnull() == True) &
                 (cfv['Activity'].str.contains('New My.TMO Order') == True) &
                 (cfv['Floodlight Attribution Type'].str.contains('Click-through') == True)), 1, 0))

        return cfv


    def f_tags(self, data):
        # Create a DataFrame of the F Tag sheet included in the Excel worksheet.
        ftags = pd.DataFrame(Range('F_Tags', 'B1').table.value, columns=Range('F_Tags', 'B1').horizontal.value)
        ftags.drop(0, inplace=True)
        ftags.drop_duplicates(subset='Expected URL', take_last=True, inplace=True)

        ftags['Expected URL'] = ftags['Expected URL'].str.replace('.html', '')
        #ftags['Expected URL'] = ftags['Expected URL'].str.replace('http://explore', 'http://www')

        # Add a new column to the DataFrame to concatenate the Group Name of the tag with the Activity Name. This will
        # give us a reference we can use to match to the tag to the data.

        data = pd.merge(data, ftags, how='left', left_on = 'Click-through URL', right_on = 'Expected URL')
        data.rename(columns={'Activity Name': 'F Tag'}, inplace = True)

        # the F Tags names that were inputted are then matched to the headers of the columns. When a match is found, the
        # reference is set in a list.
        data['Tag Name (Concatenated)'] = data['Group Name'] + " : " + data['F Tag']
        data['Tag Name (Concatenated)'].fillna('na', inplace = True)

        f_tag = []
        for i in data['Tag Name (Concatenated)']:
            for j in list(data.columns):
                tag = re.search(i, j)
                if tag:
                    f_tag.append(j)

        # After all the F Tag names have been iterated through to find the appropriate tag columns, the references are then
        # set as the intersection of the column names.
        f_tag = list(set(f_tag).intersection(data.columns))
        f_conversions = list(set(f_tag).intersection(data.columns))

        # The F Actions column that was created earlier is then updated with the sum of the F Action by row based on the
        # corresponding columns to that tag.
        data['F Actions'] = data[f_conversions].sum(axis=1)

        data.drop(['Activity ID', 'Group Name', 'Expected URL', 'Tag', 'Tag Name (Concatenated)'], axis = 1)

        return data