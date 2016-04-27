


def cfv_tab_name():
    cfv = 'CFV_Temp'

    return cfv


def sa_tab_name():
    sa = 'SA_Temp'

    return sa


def qa_tab_name():
    qa = 'Data_QA_Output'

    return qa


def add_sheets():
    sa = sa_tab_name()
    cfv = cfv_tab_name()
    qa = qa_tab_name()

    to_add = [sa, cfv, qa]

    return to_add


def delete_sheets(sheet_names):
    sa = sa_tab_name()
    cfv = cfv_tab_name()
    qa = qa_tab_name()

    sheets_to_remove = [sa, cfv, qa]

    to_delete = list(set(sheets_to_remove) & set(sheet_names))

    return to_delete
