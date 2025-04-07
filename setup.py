import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {
    "packages": ["PyQt6", "pandas", "PyPDF2", "openpyxl"],
    "excludes": ["tkinter", "unittest"],
    "include_files": [
        ("resources/icons/app_icon.ico", "resources/icons/app_icon.ico"),
    ],
    "include_msvcr": True,
}

# GUI applications require a different base on Windows
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Certificate Generator",
    version="1.0",
    description="แอปพลิเคชันสร้างเกียรติบัตร",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "main.py",
            base=base,
            icon="resources/icons/app_icon.ico",
            target_name="Certificate Generator.exe",
        )
    ],
)