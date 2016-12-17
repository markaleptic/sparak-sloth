import cx_Freeze
import sys
import os

os.environ['TCL_LIBRARY'] = "C:\\Users\\mallred\\AppData\\Local\\Continuum\\Anaconda3\\tcl\\tcl8.6"
os.environ['TK_LIBRARY']  = "C:\\Users\\mallred\\AppData\\Local\\Continuum\\Anaconda3\\tcl\\tk8.6"



base = None

if sys.platform == 'win32':
    base = "Win32GUI"

executables = [cx_Freeze.Executable("Sparak_Sloth.py", base=base, icon='slothicon.ico')]

build_packages = ["pyautogui","tkinter"]
include_files = ["slothicon.ico", "slothBackground.png", "delete_sign.png", "plus_sign.png",
                 r"C:\\Users\\mallred\\AppData\\Local\\Continuum\\Anaconda3\\DLLs\\tcl86t.dll",
                 r"C:\\Users\\mallred\\AppData\\Local\\Continuum\\Anaconda3\\DLLs\\tk86t.dll"
                ]

cx_Freeze.setup(
    name = "Sparak Accounting Sloth",
    options = {"build_exe": {"packages"     :   build_packages, 
                             "include_files":   include_files}},
    version = "1.0",
    description = "Sparak Accounting Sloth",
    executables = executables
    )
