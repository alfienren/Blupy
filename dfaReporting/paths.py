from xlwings import Range

def report_path():
    path = Range('Action_Reference', 'AG1').value

    return path


def dr_pacing_path():
    path = Range('Sheet3', 'AC1').value

    return path


def r_output_path():
    path = Range('Sheet3', 'AD1').value

    return path


def dr_pivot_path():
    path = Range('Sheet3', 'AB1').value

    return path


def client_data_path():
    save_path = str(dr_pivot_path())
    save_path = save_path[:save_path.rindex('\\')]

    return save_path