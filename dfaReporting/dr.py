from reporting import custom_variables, datafunc, categorization, floodlights, clickthroughs, report_columns, template, paths
from reporting.ddr import top_devices, forecast, dashboard
from reporting.ddr.performance import generate_emails, tables, publisher
import numpy as np
import pandas as pd
from xlwings import Range, Sheet, Workbook, Application


def raw_pivot():
    path = paths.path_select()

    ddr = pd.read_excel(path, 'data', index_cols=None, parse_cols='A:V,X,Z:AK,CR:DJ')
    ddr.fillna(0, inplace=True)

    return ddr


def dr_reporting():
    wb = Workbook.caller()
    wb.save()

    sa = datafunc.read_site_activity_report(wb.fullname, adv='dr')
    cfv2 = datafunc.read_cfv_report(wb.fullname)

    date = sa['Date'].max().strftime('%m.%d.%Y')

    feed_path = Range('Action_Reference', 'AE1').value
    excluded = top_devices.excluded_devices()

    cfv = custom_variables.custom_variable_columns(cfv2)
    cfv = custom_variables.ddr_custom_variables(cfv)

    data = sa.append(cfv)
    data = clickthroughs.strip_clickthroughs(data)

    data = custom_variables.format_custom_variable_columns(data)

    data = floodlights.a_e_traffic(data)

    data = categorization.sites(data)
    data = categorization.date_columns(data)
    data = categorization.dr_placement_message_type(data)
    data = categorization.dr_tactic(data)
    data = categorization.placement_categories(data)
    data = categorization.dr_creative_categories(data)
    data = report_columns.additional_columns(data, adv='dr')

    cfv_floodlight_columns = ['Activity', 'OrderNumber (string)', 'Plan (string)', 'Device (string)',
                              'Service (string)', 'Accessory (string)', 'Floodlight Attribution Type', 'Orders',
                              'Total GAs', 'Prepaid GAs', 'Postpaid GAs', 'Prepaid Orders', 'Postpaid Orders',
                              'Prepaid SIMs', 'Postpaid SIMs', 'Prepaid Mobile Internet',
                              'Postpaid Mobile Internet', 'Prepaid Phone', 'Postpaid Phone', 'DDR Add-a-Line', 'DDR New Devices']

    sa_columns = list(sa.columns)
    tag_columns = sa_columns[sa_columns.index('DBM Cost (USD)') + 1:]

    columns = report_columns.order_columns(adv='dr') + tag_columns + cfv_floodlight_columns

    data = data[columns]

    data.fillna(0, inplace=True)

    datafunc.merge_past_data(data, columns, wb.fullname)


    wb2 = Workbook()
    wb2.set_current()

    top_devices.top_15_devices(cfv2, feed_path, excluded)

    wb2.save(r'S:\SEA-Media\Analytics\T-Mobile\DR\Top 15 Devices Report\Top Devices Report ' + date + '.xlsx')
    wb2.close()

    wb.set_current()


def generate_dashboard():
    dashboard.generate_data()


def output_forecasts(pacing_data):
    pacing_data['Week'] = pacing_data['Date'].apply(lambda x: categorization.mondays(x))

    pacing_data = pd.pivot_table(pacing_data, index= ['Site', 'Tactic', 'Metric'],
                          columns= ['Week'], values= 'value', aggfunc= np.sum).reset_index()

    #wb = Workbook(dr_pacing_path())

    Sheet('forecast_data').clear_contents()
    Range('forecast_data', 'A1', index= False).value = pacing_data

    #wb.save()
    #wb.close()


def pacing_data_for_forecasts():
    wb = Workbook.caller()

    dr_data = raw_pivot()

    dr_forecasting = forecast.transform_dr_forecasts(dr_data)
    wb.set_current()

    Sheet('raw_pacing_data').clear_contents()
    Range('raw_pacing_data', 'A1', index=False).value = dr_forecasting

    publisher.publishers(dr_data)


def reshape_forecasts():
    Workbook.caller()

    r_data = forecast.generate_forecasts()

    pacing_data = forecast.merge_pacing_and_forecasts(r_data)

    tab = dashboard.tableau_pacing(pacing_data)

    output_forecasts(tab)


def emails_to_publishers():
    pacing_wb = Workbook.caller()
    generate_emails.generate_publisher_emails(tables.aggregated(), tables.contacts(), tables.brand_remessaging())
    tables.tables_for_emails(tables.site_tactic())

    Application(wkb=pacing_wb).xl_app.Run('Format_Tables')