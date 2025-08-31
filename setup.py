#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Setup script for creating standalone executable of Voucher Processor GUI
"""

import sys
from cx_Freeze import setup, Executable
import os

# Dependencies are automatically detected, but some modules may be missed
build_exe_options = {
    "packages": [
        "tkinter", "pandas", "openpyxl", "xml", "zipfile", 
        "tempfile", "shutil", "threading", "datetime", "pathlib"
    ],
    "excludes": ["matplotlib", "numpy", "scipy", "IPython", "jupyter"],
    "include_files": [
        ("converter.py", "converter.py"),
    ],
    "zip_include_packages": ["*"],
    "zip_exclude_packages": [],
}

# GUI applications require a different base on Windows
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="VoucherProcessor",
    version="1.0.0",
    description="Standalone Voucher Processing Application",
    options={"build_exe": build_exe_options},
    executables=[Executable("voucher_gui.py", base=base, target_name="VoucherProcessor")],
)