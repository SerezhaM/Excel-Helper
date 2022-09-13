from cx_Freeze import setup, Executable

base = None

executables = [Executable("QT_test.py", base=base)]

packages = ["pathlib"]
options = {
    'build_exe': {
        'packages':packages,
    },
}

setup(
    name = "Excel Helper",
    options = options,
    version = "0.1",
    description = '<any description>',
    executables = executables
)


# ВЫПОЛНЕНИЕ: python setup.py build !!!!!!!!!!!!!!

# APP = ['QT_test.py']
# DATA_FILES = ['Names.txt', 'main_console.py', 'Excel_Helper.py', 'Excel_Helper_names.py']
# OPTIONS = {'argv_emulation': False,
#            'semi_standalone': 'False',
#            'compressed' : 'True',
#            'packages' : ('openpyxl', 'pathlib'),
#           }

