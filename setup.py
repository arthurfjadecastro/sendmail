import sys
from cx_Freeze import setup, Executable
from tkinter import *


base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [
    Executable("main.py", base=base, icon="sr2637.ico")
]

buildOptions = dict(
    packages=[],
    includes=["random", "win32com",
              "textwrap", "pandas", "os", "tkinter", "PIL"],
    include_files=["BebasNeue-Regular.ttf", "sr2637.ico", "Empregados.xlsx", "main.spec",
                   "parabensind.jpg", "result.png", "settings.json"],
    excludes=[]
)

setup(
    name="Enviar Emails",
    version="1.0",
    description="Descrição do programa",
    options=dict(build_exe=buildOptions),
    executables=executables
)
