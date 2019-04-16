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
    version = "20.0",
    description = "Controle Cycle 2 ETD-TVX",
    # options = {"build_exe": build_exe_options},
    executables = [
    Executable("1_Insertion_DPL_OPTIMUM.py"),
    Executable("2_MaJ_DPL_OPTIMUM_NEG-Infos.py"),
    Executable("3_cycle_2_generation_NEGO.py"),
    Executable("4_MaJ_DPL_OPTIMUM_ETD-Avc.py"),
    Executable("5_MaJ_SPI.py"),
    Executable("6_MaJ_DPL_OPTIMUM_ETD-FCI-EZA.py"),
    Executable("7_cycle_2_generation_ETD-TVX.py")
    ])