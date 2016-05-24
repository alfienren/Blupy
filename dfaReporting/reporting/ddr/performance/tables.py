import pandas as pd
from xlwings import Range

from reporting import categorization
from reporting.constants import DrPerformance


def contacts():
    contacts_df = pd.DataFrame(Range('Publisher_Contacts', 'A1').table.value,
                            columns= Range('Publisher_Contacts', 'A1').horizontal.value)
    contacts_df.drop(0, inplace= True)

    return contacts_df


def aggregated():
    pubs_combined = pd.DataFrame(Range(DrPerformance.pub_performance_sheet,
                                       DrPerformance.aggregate_column + str(DrPerformance.row_num)).table.value,
                                 columns=Range(DrPerformance.pub_performance_sheet,
                                               DrPerformance.aggregate_column + str(
                                                   DrPerformance.row_num)).horizontal.value)
    pubs_combined.drop(0, inplace=True)

    return pubs_combined


def brand_remessaging():
    br = pd.DataFrame(Range(DrPerformance.pub_performance_sheet, DrPerformance.br_column + str(
        DrPerformance.row_num)).table.value, columns=Range(DrPerformance.pub_performance_sheet,
                                                           DrPerformance.br_column + str(
                                                               DrPerformance.row_num)).horizontal.value)
    br.drop(0, inplace=True)

    return br


def site_tactic():
    pub_performance = pd.DataFrame(Range(DrPerformance.pub_performance_sheet, DrPerformance.site_tactic_column + str(
        DrPerformance.row_num)).table.value, columns=Range(DrPerformance.pub_performance_sheet,
                                                           DrPerformance.site_tactic_column + str(
                                                               DrPerformance.row_num)).horizontal.value)
    pub_performance.drop(0, inplace=True)

    pub_performance['CPO'] = pub_performance['Spend'].astype(float) / pub_performance['Orders'].astype(float)

    pub_performance.rename(columns= {'Orders':'Total ' + categorization.qquarter() + ' Orders',
                                     'Spend':'Total ' + categorization.qquarter() + ' Spend',
                                     'Placement Messaging Type':'Tactic'}, inplace= True)

    return pub_performance


def tables_for_emails(pub_performance):
    Range(DrPerformance.pub_performance_sheet, 'A30:A1000').clear_contents()

    site_list = list(pub_performance['Site'].unique())
    week = Range(DrPerformance.pub_performance_sheet, 'D' + str(
        DrPerformance.row_num - 1)).value

    start_column = 'B'
    site_column = 'A'

    cell_number = Range(DrPerformance.pub_performance_sheet, 'A13').vertical.last_cell.row + 3

    for i in site_list:
        df = pub_performance[pub_performance['Site'] == i]
        df = df[['Tactic', 'Total ' + categorization.qquarter() + ' Orders',
                 'Total ' + categorization.qquarter() + ' Spend', 'CPO', 'Q2 CPO Goal']]
        Range(DrPerformance.pub_performance_sheet, start_column + str(cell_number), index = False).value = df
        Range(DrPerformance.pub_performance_sheet, site_column + str(cell_number), index = False).value = i
        cell_number += 7
