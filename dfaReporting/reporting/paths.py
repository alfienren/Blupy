import os
from Tkinter import Tk
from tkFileDialog import askopenfilename


def path_select():
    Tk().withdraw()

    file_path = askopenfilename()
    file_path = os.path.normpath(file_path)

    return file_path
