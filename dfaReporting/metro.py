from xlwings import Workbook
from reporting import *


def generate_metro_reporting():
    wb = Workbook.caller()
    wb.save()

    sa = datafunc.read_site_activity_report(wb.fullname, adv='metro')
    cfv = datafunc.read_cfv_report(wb.fullname)

    data = sa.append(cfv)
    data = floodlights.a_e_traffic(data, adv='metro')
    data = categorization.date_columns(data)
    data = categorization.categorize_report(data, adv='metro')

    data = report_columns.additional_columns(data, adv='metro')

    sa_columns = list(sa.columns)
    tag_columns = sa_columns[sa_columns.index('Clicks') + 1:]

    columns = report_columns.order_columns(adv='metro') + tag_columns

    data = data[columns]
    data = data.fillna(0)

    datafunc.merge_past_data(data, columns, wb.fullname)

    qa.placement_qa(data)