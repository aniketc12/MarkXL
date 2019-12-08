from cx_Freeze import setup, Executable


setup(name = "MarkXL" ,
      version = "0.1" ,
      description = "" ,
      options = {"build_exe": {"packages": ["tkinter", "openpyxl", "shelve", "dbm", "os"]}},
      executables = [Executable("MarkXL.py", base = "Win32GUI")])
