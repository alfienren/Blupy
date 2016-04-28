import pandas as pd
from xlwings import Range


def performance_sheet():
    pub_performance_sheet = 'Publisher Performance'

    return pub_performance_sheet


def qquarter():
    quarter = 'Q2'

    return quarter


def quarter_start_year(start='quarter_start'):
    if start != 'quarter_start':
        quarter = '4/18/2016'
    else:
        quarter = '4/1/2016'

    return quarter


def week_of(dr):
    week = 'Spend Week of ' + str(dr['Week'].max().strftime('%m/%d/%Y').lstrip('0').replace(' 0', ' '))

    return week


def row():
    row_num = 13

    return row_num


def br_column():
    column = 'L'

    return column


def site_tactic_column():
    column = 'A'

    return column


def aggregated_column():
    column = 'H'

    return column
