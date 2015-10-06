from xlwings import Range
import re

def actions(data):

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