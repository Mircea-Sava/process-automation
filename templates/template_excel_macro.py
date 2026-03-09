from process_automation import run_excel_macro

FILE = r"Z:\path\to\your_workbook.xlsm"                                      # <-- CHANGE: full path to your .xlsm file

macros = [                                                           # <-- ADD: one (macro_name, module_name) per line
    ("MacroName1", "ModuleName1"),                                         # <-- CHANGE: ("MacroName", "ModuleName")
    ("MacroName2", "ModuleName2"),                                         # <-- CHANGE: ("MacroName", "ModuleName")
]

for macro_name, module_name in macros:
    run_excel_macro(file_path=FILE, macro_name=macro_name, module_name=module_name, visible=False)  # <-- CHANGE: False to run in background
