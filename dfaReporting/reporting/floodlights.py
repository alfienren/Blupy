import re

from xlwings import Range
import pandas as pd


def a_e_traffic(data, adv='tmo'):
    # Set the data column names to a variable
    column_names = data.columns

    a_actions = Range('Action_Reference', 'A2').vertical.value
    b_actions = Range('Action_Reference', 'B2').vertical.value
    c_actions = Range('Action_Reference', 'C2').vertical.value
    d_actions = Range('Action_Reference', 'D2').vertical.value
    e_actions = Range('Action_Reference', 'E2').vertical.value

    # For each action tag category (A - E), search the column names to find the action tag. A new list of the A - E
    # actions are compiled from the matches in the column names.
    a_actions = list(set(a_actions).intersection(column_names))
    b_actions = list(set(b_actions).intersection(column_names))
    c_actions = list(set(c_actions).intersection(column_names))
    d_actions = list(set(d_actions).intersection(column_names))
    e_actions = list(set(e_actions).intersection(column_names))

    if adv != 'tmo':
        language = '|'.join(list(['Spanish', 'Hispanic', 'SL']))

        lat_a, lat_b, lat_c, lat_d, lat_e, gm_a, gm_b, gm_c, gm_d, gm_e = \
        [], [], [], [], [], [], [], [], [], []

        for i in a_actions:
            lat = re.search(language, i)
            if lat:
                lat_a.append(i)
            if not lat:
                gm_a.append(i)

        for i in b_actions:
            lat = re.search(language, i)
            if lat:
                lat_b.append(i)
            if not lat:
                gm_b.append(i)

        for i in c_actions:
            lat = re.search(language, i)
            if lat:
                lat_c.append(i)
            if not lat:
                gm_c.append(i)

        for i in d_actions:
            lat = re.search(language, i)
            if lat:
                lat_d.append(i)
            if not lat:
                gm_d.append(i)

        for i in e_actions:
            lat = re.search(language, i)
            if lat:
                lat_e.append(i)
            if not lat:
                gm_e.append(i)

        gm_a = list(set(gm_a).intersection(column_names))
        gm_b = list(set(gm_b).intersection(column_names))
        gm_c = list(set(gm_c).intersection(column_names))
        gm_d = list(set(gm_d).intersection(column_names))
        gm_e = list(set(gm_e).intersection(column_names))

        lat_a = list(set(lat_a).intersection(column_names))
        lat_b = list(set(lat_b).intersection(column_names))
        lat_c = list(set(lat_c).intersection(column_names))
        lat_d = list(set(lat_d).intersection(column_names))
        lat_e = list(set(lat_e).intersection(column_names))

        data['GM A Actions'] = data[gm_a].sum(axis=1)
        data['GM B Actions'] = data[gm_b].sum(axis=1)
        data['GM C Actions'] = data[gm_c].sum(axis=1)
        data['GM D Actions'] = data[gm_d].sum(axis=1)
        data['GM E Actions'] = data[gm_e].sum(axis=1)

        data['Hispanic A Actions'] = data[lat_a].sum(axis=1)
        data['Hispanic B Actions'] = data[lat_b].sum(axis=1)
        data['Hispanic C Actions'] = data[lat_c].sum(axis=1)
        data['Hispanic D Actions'] = data[lat_d].sum(axis=1)
        data['Hispanic E Actions'] = data[lat_e].sum(axis=1)

        data['Total A Actions'] = data['GM A Actions'] + data['Hispanic A Actions']
        data['Total B Actions'] = data['GM B Actions'] + data['Hispanic B Actions']
        data['Total C Actions'] = data['GM C Actions'] + data['Hispanic C Actions']
        data['Total D Actions'] = data['GM D Actions'] + data['Hispanic D Actions']
        data['Total E Actions'] = data['GM E Actions'] + data['Hispanic E Actions']

        data['Orders'] = data['Transaction Count']

    # With the references set for each action tag category, sum the tags along rows and create columns for each action
    # tag bucket.
    data['A Actions'] = data[a_actions].sum(axis=1)
    data['B Actions'] = data[b_actions].sum(axis=1)
    data['C Actions'] = data[c_actions].sum(axis=1)
    data['D Actions'] = data[d_actions].sum(axis=1)
    data['E Actions'] = data[e_actions].sum(axis=1)

    # Similar to the logic to get the sum of action tags for each respective category, the following finds the sum
    # of post-click and impression activity as well as Store Locator Visits.

    # For post-view and click activity, the column names of the data are searched for columns containing 'View-through',
    # 'Click-through' or 'Store Locator'. Column matches are then stored in lists.
    view_through = []
    for item in column_names:
        view = re.search('View-through Conversions', item)
        if view:
            view_through.append(item)

    click_through = []
    for item in column_names:
        click = re.search('Click-through Conversions', item)
        if click:
            click_through.append(item)

    store_locator = []
    for item in column_names:
        locator = re.search('Store Locator 2', item)
        if locator:
            store_locator.append(item)

    # The matches against the column names are then set
    view_based = list(set(view_through).intersection(column_names))
    click_based = list(set(click_through).intersection(column_names))
    SLV_conversions = list(set(store_locator).intersection(column_names))

    # With the matching references set, sum the matches for post-click and impression activity as well as SLV by row,
    # creating columns for each.
    data['Post-Click Activity'] = data[click_based].sum(axis=1)
    data['Post-Impression Activity'] = data[view_based].sum(axis=1)
    data['Store Locator Visits'] = data[SLV_conversions].sum(axis=1)

    # Create columns for Awareness and Consideration actions. Awareness Actions are the sum of A and B actions,
    # Consideration is the sum of C and D actions
    # Traffic Actions are the total of Awareness and Consideration Actions, or A - D actions.
    data['Awareness Actions'] = data['A Actions'] + data['B Actions']
    data['Consideration Actions'] = data['C Actions'] + data['D Actions']
    data['Traffic Actions'] = data['Awareness Actions'] + data['Consideration Actions']

    return data


def f_tags(data):

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
