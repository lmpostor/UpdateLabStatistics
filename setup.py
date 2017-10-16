import sys
from cx_Freeze import setup, Executable

#used to build this into an exe to run on machines without python. Only will work on 64bit Windows

setup(
    name = "HrlyStationStats",
    version = "1.0",
    description = "Send hourly station stats and update google spreadsheet",
    executables = [Executable("updatestats.py", base = "Win32GUI")])
