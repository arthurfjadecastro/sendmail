import sys
from cx_Freeze import setup, Executable
from cProfile import label
from random import randint
from turtle import back
import win32com.client as client
import textwrap
import pandas as pd
import os
from tkinter.messagebox import showinfo
from tkinter import *
from PIL import Image, ImageFont, ImageDraw


base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [
        Executable("main.py", base=base,ico="sr2637.ico")
]

buildOptions = dict(
        packages = [],
        includes = ["cProfile","random","turtle","win32com","textwrap","pandas","os","tkinter","PIL"],
        include_files = ["BebasNeue-Regular.ttf", "chef.ico","Empregados.xlsx","main.spec","parabensind.jpg","result.png","settings.json","sr2637.ico","sr2637.png"],
        excludes = []
)

setup(
    name = "Enviar Emails",
    version = "1.0",
    description = "Descrição do programa",
    options = dict(build_exe = buildOptions),
    executables = executables
 )
