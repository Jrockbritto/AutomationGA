from cx_Freeze import setup, Executable

base = None


executables = [Executable("automationcol.py",
               base=base,
               icon="img/icon.ico",
               targetName="Automation"
               )]

packages = ['os','sys', 'logging']

options = {
    'build_exe': {    
        'packages':packages,
        'include_files': ['img/'],
        'includes': ["selenium", "webdriver_manager.chrome", "math","openpyxl", "datetime"],
        'build_exe': 'Automation',
        'excludes': ['tkinter', 'ctypes', 'html', 'pydoc_data', 'test', 'xmlrpc'] #otimizando script
    },
}

setup(
    name = "AutomaçãoGA",
    options = options,
    version = "1.0",
    description = 'Faz um banco com os links e algumas info de todos os elementos do site.',
    executables = executables
)
