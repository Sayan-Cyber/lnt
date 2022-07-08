import cx_Freeze
import sys
import os
from cx_Freeze.dist import Distribution
base = None
if sys.platform == 'win32':
    base = "Win32GUI"


os.environ['TCL_LIBRARY'] = r"C:\\Users\\20323801\\AppData\\Local\\Programs\Python\\Python310\\tcl\\tcl8.6"
os.environ['TK_LIBRARY'] = r"C:\\Users\\20323801\\AppData\\Local\\Programs\\Python\\Python310\\tcl\\tk8.6"





executables = [cx_Freeze.Executable("Main2.py", base=base,
                                    copyright='SC', shortcut_name='Pragati', shortcut_dir='DesktopFolder')]

cx_Freeze.setup(name="ProjectII)", options={"build_exe": {"packages": ["os", "sys", "openpyxl",
                                                                                       'PIL', 'pyglet', 'webbrowser'], "optimize": 2, "include_files": ['Additional_File']}}, version="7.8.22", author='L&T Construction(WET IC)', description="Larsen & Toubro ", executables=executables)
