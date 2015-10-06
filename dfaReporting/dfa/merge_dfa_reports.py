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

    data = floodlights.custom_floodlight_tags(data)
    data = floodlights.action_tags(data)

    data = categorization.placements(data)
    data = categorization.sites(data)
    data = categorization.creative(data)
    data = categorization.language(data)
    data = categorization.date_columns(data)

    data = floodlights.f_tags(data)

    return data