from reporting.ddr.performance.performance import *
from reporting.ddr.performance.performance import performance_sheet, qquarter
from xlwings import Range


def contacts():
    contacts = pd.DataFrame(Range('Publisher_Contacts', 'A1').table.value,
                            columns= Range('Publisher_Contacts', 'A1').horizontal.value)
    contacts.drop(0, inplace= True)

    return contacts


def aggregated():
    pubs_combined = pd.DataFrame(Range(performance_sheet(), 'A' + str(row())).table.value,
                                 columns= Range(performance_sheet(), 'A' + str(row())).horizontal.value)
    pubs_combined.drop(0, inplace= True)

    return pubs_combined


def brand_remessaging():
    br = pd.DataFrame(Range(performance_sheet(), 'A' + str(row() - 1)
                            + ':E' + str(row() - 1)).vertical.value,
                      columns= Range(performance_sheet(), 'A' + str(row() - 1)
                            + ':E' + str(row() - 1)).value)
    br.drop(0, inplace=True)

    return br


def site_tactic():
    pub_performance = pd.DataFrame(Range(performance_sheet(), 'A' + str(row() - 1)
                                         + ':E' + str(row() - 1)).vertical.value,
                                   columns= Range(performance_sheet(), 'A' + str(row() - 1)
                                         + ':E' + str(row() - 1)).value)
    pub_performance.drop(0, inplace= True)

    pub_performance['CPO'] = pub_performance['Spend'].astype(float) / pub_performance['Orders'].astype(float)

    pub_performance.rename(columns= {'Orders':'Total ' + qquarter() + ' Orders',
                                     'Spend':'Total ' + qquarter() + ' Spend',
                                     'Placement Messaging Type':'Tactic'}, inplace= True)

    pub_performance = goals(pub_performance)

    return pub_performance


def goals(data):
    goals = pd.DataFrame(Range('Goals', 'A1').table.value,
                            columns=Range('Goals', 'A1').horizontal.value)

    goals.drop(0, inplace=True)

    return goals


def tables_for_emails(pub_performance):
    Range(performance_sheet(), 'A' + str(row() + 2) + ':A1000').clear_contents()

    site_list = list(pub_performance['Site'].unique())
    week = Range(performance_sheet(), 'D' + str(row() - 1)).value

    start_column = 'B'
    site_column = 'A'

    cell_number = Range(performance_sheet(), 'A13').vertical.value.cell_number.last_cell.row
    #cell_number = cell_number.last_cell.row
    cell_number = cell_number.offset(row_offset = 4)

    for i in site_list:
        df = pub_performance[pub_performance['Site'] == i]
        df = df[['Tactic', 'Total ' + qquarter() + ' Orders', week,
                 'Total ' + qquarter() + ' Spend', 'CPO', qquarter() + ' CPO Goal']]
        Range(performance_sheet(), start_column + str(cell_number), index = False).value = df
        Range(performance_sheet(), site_column + str(cell_number), index = False).value = i
        cell_number += 7