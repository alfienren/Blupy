import pandas as pd
import re
from xlwings import Range

def f_tags(data):
    
    # Create a DataFrame of the F Tag sheet included in the Excel worksheet.
    ftags = pd.DataFrame(Range('F_Tags', 'B1').table.value, columns=Range('F_Tags', 'B1').horizontal.value)
    ftags.drop(0, inplace=True)

    ftags['Expected URL'] = ftags['Expected URL'].str.replace('.html', '')
    ftags['Expected URL'] = ftags['Expected URL'].str.replace('http://explore', 'http://www')

    # Add a new column to the DataFrame to concatenate the Group Name of the tag with the Activity Name. This will
    # give us a reference we can use to match to the tag to the data.

    data = pd.merge(data, ftags, how='left', left_on = 'Click-through URL', right_on = 'Expected URL')
    data.rename(columns={'Activity Name': 'F Tag'}, inplace = True)

    # for each cell in the range, an INDEX + MATCH formula is entered to find the F Tag for the URL listed in the column
    # to the left.
    '''
    for cell in Range('working', 'F2').vertical:
        url = cell.offset(0, -1).get_address(False, False, False)
        cell.formula = '=IFERROR(INDEX(F_Tags!G:G,MATCH(working!' + url + ',F_Tags!E:E,0)),"na")'
    '''
    # With the F Tag names entered, update the DataFrame's F Tag column with the inputted data.
    #data['F Tag'] = Range('working', 'F2').vertical.value

    # the F Tags names that were inputted are then matched to the headers of the columns. When a match is found, the
    # reference is set in a list.
    ftags['Tag Name (Concatenated)'] = ftags['Group Name'] + " : " + ftags['Activity Name']

    f_tags = []
    for i in data['Tag Name (Concatenated)']:
        for j in data.columns:
            tag = re.search(i, j)
            if tag:
                f_tags.append(j)

    # After all the F Tag names have been iterated through to find the appropriate tag columns, the references are then
    # set as the intersection of the column names.
    f_tags = list(set(f_tags).intersection(data.columns))
    f_conversions = list(set(f_tags).intersection(data.columns))

    # The F Actions column that was created earlier is then updated with the sum of the F Action by row based on the
    # corresponding columns to that tag.
    data['F Actions'] = data[f_conversions].sum(axis=1)

    data.drop(['Activity ID', 'Group Name', 'Expected URL', 'Tag', 'Tag Name (Concatenated)'], axis = 1)

    return data