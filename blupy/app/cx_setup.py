import sys
from cx_Freeze import setup, Executable

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

options = {
    'build_exe': {
        'packages': ['pandas'],
        'includes': ['atexit', 'PySide.QtNetwork', 'pandas']
    }
}

executables = [
    Executable('app.py', base=base, icon='lion.ico')
]

setup(name='Blupy',
      version='0.1',
      description='Blupy app',
      options=options,
      executables=executables
      )