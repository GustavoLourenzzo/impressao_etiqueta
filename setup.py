#setup.py
import sys,os
from cx_Freeze import setup, Executable

os.environ['TCL_LIBRARY'] = r'C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\tcl\tk8.6'
setup(
    name = "Impressor de Etiqueta",
    version = "1.0.1",
    description="Impressor de Etiqueta : Impressao de etiqueta termica",
    options = {"build_exe": {
        'packages': ["wx","win32ui","win32print","win32con","win32gui","os","sys","configparser","shutil","PIL","idna"],
        #'include_files': include_files,
        'include_msvcr': True,
       },
       
    },
    executables = [Executable(r"C:\Users\Gustavo\Desktop\Prog. Impressao\ProgCorrigido\main.py")]
    )


