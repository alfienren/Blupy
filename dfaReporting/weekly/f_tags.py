__author__ = 'aarschle1'

import pandas as pd
import re
from xlwings import Range

def f_tags():
    
    # Create a DataFrame of the F Tag sheet included in the Excel worksheet.
    ftags = pd.DataFrame(Range('F_Tags', 'B1').table.value, columns=Range('F_Tags', 'B1').horizontal.value)
    ftags.drop(0, inplace=True)

    # Add a new column to the DataFrame to concatenate the Group Name of the tag with the Activity Name. This will
    # give us a reference we can use to match to the tag to the data.
    ftags['Tag Name (Concatenated)'] = ftags['Group Name'] + " : " + ftags['Activity Name']
    Range('F_Tags', 'G2', index=False).value = ftags['Tag Name (Concatenated)']

    # The F Tag Range is set as column F in the working data (The F Tag column)

    #f_tag_range = Range('working', 'F2').vertical

    # for each cell in the range, an INDEX + MATCH formula is entered to find the F Tag for the URL listed in the column
    # to the left.
    for cell in Range('working', 'F2').vertical:
        url = cell.offset(0, -1).get_address(False, False, False)
        cell.formula = '=IFERROR(INDEX(F_Tags!G:G,MATCH(working!' + url + ',F_Tags!E:E,0)),"na")'

    # With the F Tag names entered, update the DataFrame's F Tag column with the inputted data.
    #data['F Tag'] = Range('working', 'F2').vertical.value

    # the F Tags names that were inputted are then matched to the headers of the columns. When a match is found, the
    # reference is set in a list.
    f_tags = []
    for i in Range('working', 'F2').vertical:
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

    # Strip everything before the colon in the F Tag column to remove the group name.
    data['F Tag'] = data['F Tag'].apply(lambda x: str(x).split(':')[-1])