import cx_Freeze
import sys
import os
from cx_Freeze.dist import Distribution
base = None
if sys.platform == 'win32':
    base = "Win32GUI"




os.environ['TCL_LIBRARY'] = r"C:\Users\20323801\AppData\Local\Programs\Python\Python39\tcl\tcl8.6"
os.environ['TK_LIBRARY'] = r"C:\Users\20323801\AppData\Local\Programs\Python\Python39\tcl\tk8.6"


bdist_msi_options = {
    'install_icon': "Additional_File\icons\icon.ico"}


executables = [cx_Freeze.Executable("Main2.py", base=base,icon="Additional_File\icons\icon.ico",copyright='SC', shortcut_name='Pragati', shortcut_dir='DesktopFolder')]

cx_Freeze.setup( name = "ProjectII)",options = {"build_exe": {"packages":["os","sys","docxtpl",
                                                                                       'PIL','pyglet','webbrowser'],"optimize":2,"include_files":['Additional_File'],
                                                                           'excludes':['pygments','debugpy','zmq','pandas','matplotlib','numpy',
                                                                            'tkinterdnd2','docutils','docxpdf','decorator','image','importlib_metadata','keyboard','pyautogui',
                                                                          'pycparser','pymsgbox','pyperclip','requests','sqlparse','tkinterdnd']},
                                                             'bdist_msi': bdist_msi_options},version = "31.3.22",author= 'L&T Construction(WET IC)', description =  "Larsen & Toubro ",executables = executables)
