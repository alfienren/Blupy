import re
from xlwings import Range
import numpy as np
import pandas as pd


def custom_floodlight_tags(data):

    # CFV columns for Plans, Services, etc. that were created earlier have blank values replaced with 0.
    data['Plans'].fillna(0, inplace=True)
    data['Services'].fillna(0, inplace=True)
    data['Devices'].fillna(0, inplace=True)
    data['Accessories'].fillna(0, inplace=True)
    data['Orders'].fillna(0, inplace=True)

    # If the count of plans, services, accessories, devices, or orders is less than 1, the string is set to blank. If
    # the count is 1 or greater, the string is associated to the count.
    data['Plan (string)'] = np.where(data['Plans'] < 1, '', data['Plan (string)'])
    data['Service (string)'] = np.where(data['Services'] < 1, '', data['Service (string)'])
    data['Accessory (string)'] = np.where(data['Accessories'] < 1, '', data['Accessory (string)'])
    data['Device (string)'] = np.where(data['Devices'] < 1, "", data['Device (string)'])
    data['OrderNumber (string)'] = np.where(data['Orders'] < 1, '', data['OrderNumber (string)'])
    data['Activity'] = np.where(data['Orders'] < 1, '', data['Activity'])
    data['Floodlight Attribution Type'] = np.where(data['Orders'] < 1, '', data['Floodlight Attribution Type'])
    data['Devices'] = np.where(data['Device (string)'].str.contains('nan') == True, 0, data['Devices'])

    return data

def action_tags(data):

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

def run_action_floodlight_tags(data):

    data = action_tags(data)
    data = custom_floodlight_tags(data)

    return data