from xlwings import Range


class TabNames:
    site_activity = 'SA_Temp'
    floodlight_variable = 'CFV_Temp'
    qa_tab_name = 'Data_QA_Output'
    search_output = 'Search_Output'


class DrPerformance:
    pub_performance_sheet = 'Publisher Performance'
    row_num = 13
    br_column = 'L'
    site_tactic_column = 'A'
    aggregate_column = 'H'


class StaticPaths:
    offline = Range('Ref', 'A3').value
    online = Range('Ref', 'A4').value
    social = Range('Ref', 'A5').value