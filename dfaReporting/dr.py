from xlwings import Range, Workbook, Application

from reporting import custom_variables, datafunc, categorization, floodlights, clickthroughs, report_columns
from reporting.ddr import devices, dashboard
from reporting.ddr.performance import generate_emails, tables


def dr_reporting():
    wb = Workbook.caller()
    wb.save()

    sa = datafunc.read_site_activity_report(wb.fullname, adv='dr')
    cfv2 = datafunc.read_cfv_report(wb.fullname)

    date = sa['Date'].max().strftime('%m.%d.%Y')

    feed_path = Range('Action_Reference', 'AE1').value
    excluded = devices.excluded_devices()

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

    devices.top_15_devices(cfv2, feed_path, excluded)

    wb2.save(r'S:\SEA-Media\Analytics\T-Mobile\DR\Top 15 Devices Report\Top Devices Report ' + date + '.xlsx')
    wb2.close()

    wb.set_current()


def generate_dashboard():
    dashboard.generate_data()


def emails_to_publishers():
    pacing_wb = Workbook.caller()
    generate_emails.generate_publisher_emails(tables.aggregated(), tables.contacts(), tables.brand_remessaging())
    tables.tables_for_emails(tables.site_tactic())

    Application(wkb=pacing_wb).xl_app.Run('Format_Tables')