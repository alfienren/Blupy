from xlwings import Range


def dr_pacing_path():
    path = Range('Sheet3', 'AC1').value

    return path


def r_output_path():
    path = Range('Sheet3', 'AD1').value

    return path


def dr_pivot_path():
    path = Range('Sheet3', 'AB1').value

    return path
