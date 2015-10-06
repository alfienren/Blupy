from xlwings import Workbook
from dfa import *

def create_report():

    wb = Workbook.caller()

    wb.save()

    sa = load_dfa_reports.raw_sa()

    cfv = load_dfa_reports.raw_cfv()

    cfv = cfv_report.cfv_data(cfv)
    cfv = cfv_report.get_creative_field(cfv)

    data = sa.append(cfv)

    data = clickthroughs.strip_clickthroughs(data)

    data = floodlights.floodlight_data(data)
    data = action_tags.actions(data)

    data = categorization.placements(data)
    data = categorization.sites(data)
    data = categorization.creative(data)

    data = f_tags.f_tags(data)
    data = data_output.output(data)

    columns = data_output.columns(sa)
    data = data[columns]
    data.fillna(0, inplace=True)