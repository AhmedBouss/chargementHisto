import sys
from cx_Freeze import setup, Executable

# # Dependencies are automatically detected, but it might need fine tuning.
# build_exe_options = {"packages": ["os"], "excludes": ["tkinter"]}

# # GUI applications require a different base on Windows (the default is for a
# # console application).
# base = None
# if sys.platform == "win32":
#     base = "Win32GUI"


setup(
	name = "Cycle_2_NEG_ETD_TVX",
    version = "1.0",
    description = "Controle Cycle 2 ETD-TVX",
    # options = {"build_exe": build_exe_options},
    executables = [
    Executable("1_Generation_DPL_OPTIMUM.py"),
    ])