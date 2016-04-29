import pandas as pd
from xlwings import Range

import reporting.ddr.performance.common


def contacts():
    contacts_df = pd.DataFrame(Range('Publisher_Contacts', 'A1').table.value,
                            columns= Range('Publisher_Contacts', 'A1').horizontal.value)
    contacts_df.drop(0, inplace= True)

    return contacts_df


def aggregated():
    pubs_combined = pd.DataFrame(Range(reporting.ddr.performance.common.performance_sheet(), reporting.ddr.performance.common.aggregated_column() + str(
        reporting.ddr.performance.common.row())).table.value,
                                 columns= Range(reporting.ddr.performance.common.performance_sheet(), reporting.ddr.performance.common.aggregated_column() + str(
                                     reporting.ddr.performance.common.row())).horizontal.value)
    pubs_combined.drop(0, inplace= True)

    return pubs_combined


def brand_remessaging():
    br = pd.DataFrame(Range(reporting.ddr.performance.common.performance_sheet(), reporting.ddr.performance.common.br_column() + str(
        reporting.ddr.performance.common.row())).table.value,
                      columns= Range(reporting.ddr.performance.common.performance_sheet(), reporting.ddr.performance.common.br_column()
                                     + str(reporting.ddr.performance.common.row())).horizontal.value)
    br.drop(0, inplace=True)

    return br


def site_tactic():
    pub_performance = pd.DataFrame(Range(reporting.ddr.performance.common.performance_sheet(), reporting.ddr.performance.common.site_tactic_column() + str(
        reporting.ddr.performance.common.row())).table.value,
                      columns= Range(reporting.ddr.performance.common.performance_sheet(), reporting.ddr.performance.common.site_tactic_column()
                                     + str(reporting.ddr.performance.common.row())).horizontal.value)
    pub_performance.drop(0, inplace= True)

    pub_performance['CPO'] = pub_performance['Spend'].astype(float) / pub_performance['Orders'].astype(float)

    pub_performance.rename(columns= {'Orders':'Total ' + reporting.ddr.performance.common.qquarter() + ' Orders',
                                     'Spend':'Total ' + reporting.ddr.performance.common.qquarter() + ' Spend',
                                     'Placement Messaging Type':'Tactic'}, inplace= True)

    #pub_performance = reporting.ddr.performance.common.goals()

    return pub_performance


def tables_for_emails(pub_performance):
    Range(reporting.ddr.performance.common.performance_sheet(), 'A30:A1000').clear_contents()

    site_list = list(pub_performance['Site'].unique())
    week = Range(reporting.ddr.performance.common.performance_sheet(), 'D' + str(
        reporting.ddr.performance.common.row() - 1)).value

    start_column = 'B'
    site_column = 'A'

    cell_number = Range(reporting.ddr.performance.common.performance_sheet(), 'A13').vertical.last_cell.row + 3

    for i in site_list:
        df = pub_performance[pub_performance['Site'] == i]
        df = df[['Tactic', 'Total ' + reporting.ddr.performance.common.qquarter() + ' Orders',
                 'Total ' + reporting.ddr.performance.common.qquarter() + ' Spend', 'CPO', 'Q2 CPO Goal']]
        Range(reporting.ddr.performance.common.performance_sheet(), start_column + str(cell_number), index = False).value = df
        Range(reporting.ddr.performance.common.performance_sheet(), site_column + str(cell_number), index = False).value = i
        cell_number += 7
