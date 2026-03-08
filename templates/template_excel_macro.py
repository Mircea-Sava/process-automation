import sys
from pathlib import Path
sys.path.append(str(Path(__file__).resolve().parent / "functions"))
from excel_macro import run_excel_macro

FILE = r"Z:\path\to\your_workbook.xlsm"                              # <-- CHANGE: full path to your .xlsm file

macros = [                                                           # <-- ADD: one (macro_name, module_name) per line
    ("MyMacro1", "Module1"),                                         # <-- CHANGE: ("MacroName", "ModuleName")
    ("MyMacro2", "Module2"),                                         # <-- CHANGE: ("MacroName", "ModuleName")
]

for macro_name, module_name in macros:
    run_excel_macro(file_path=FILE, macro_name=macro_name, module_name=module_name)
