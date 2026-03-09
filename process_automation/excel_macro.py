# Excel Macro Runner
# ==================
# Opens an Excel file (even if already open) and runs a VBA macro by name.
#
# Usage:
#   from process_automation import run_excel_macro
#
#   run_excel_macro(
#       file_path=r"Z:\some_folder\my_workbook.xlsm",
#       macro_name="MyMacroName",
#       module_name="Module1"       # optional — prefix with module name
#   )

import win32com.client
import os


def run_excel_macro(file_path: str, macro_name: str, module_name: str = None, visible: bool = True):
    """Open an Excel file (or attach to it if already open) and run a macro by name."""
    file_path = os.path.abspath(file_path)
    filename = os.path.basename(file_path)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = visible

    # Check if the workbook is already open
    wb = None
    for i in range(excel.Workbooks.Count):
        if excel.Workbooks(i + 1).Name.lower() == filename.lower():
            wb = excel.Workbooks(i + 1)
            break

    if wb is None:
        wb = excel.Workbooks.Open(file_path)

    # Run the macro
    if module_name:
        macro_ref = f"'{wb.Name}'!{module_name}.{macro_name}"
    else:
        macro_ref = f"'{wb.Name}'!{macro_name}"

    excel.Application.Run(macro_ref)
    print(f"Macro '{macro_name}' executed successfully.")
