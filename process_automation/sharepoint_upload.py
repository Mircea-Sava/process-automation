# SharePoint Excel Uploader
# =========================
# Uploads DataFrames to SharePoint as Excel files using a template-based approach.
#
# How it works:
#   1. Uses the template .xlsx file you provide via template_path
#   2. Copies the template to Desktop as a working copy (renamed to your output filename)
#   3. Opens the working copy with xlwings and writes your DataFrame into the first sheet
#   4. Saves and copies the file to the SharePoint folder
#   5. Deletes the Desktop working copy (unless keep_desktop_copy=True)
#
# Usage:
#   from process_automation import save_excel_to_sharepoint
#
#   result = save_excel_to_sharepoint(df,
#       template_path=r"Z:\path\to\TEMPLATE_DO_NOT_DELETE.xlsx",
#       sharepoint_folder=r"Z:\path\to\your_sharepoint_folder",
#       output_filename_prefix="MyReport_2026-02-26",
#   )

import os
import shutil
import stat
import datetime as _dt
from typing import Dict, Optional
import pandas as pd
import xlwings as xw

_TEMPLATE_PATTERN = "TEMPLATE_DO_NOT_DELETE"

_EXCEL_FORMATS = {
    "number":     "0.00",
    "currency":   "$#,##0.00",
    "date":       "YYYY-MM-DD",
    "time":       "HH:MM:SS",
    "percentage": "0.00%",
    "fraction":   "# ?/?",
    "text":       "@",
    "general":    "General",
}


def save_excel_to_sharepoint(
    df: pd.DataFrame,
    template_path: str,
    sharepoint_folder: str,
    output_filename_prefix: str,
    keep_desktop_copy: bool = False,
    excel_visible: bool = False,
    column_types: Optional[Dict[str, str]] = None
) -> Dict[str, str]:

    desktop_folder = os.path.join(os.environ["USERPROFILE"], "Desktop")
    try:
        os.makedirs(sharepoint_folder, exist_ok=True)
    except OSError as e:
        print(f"   [Warning] Could not create SharePoint folder dynamically. Relying on existing path. ({e})")
    os.makedirs(desktop_folder, exist_ok=True)

    # Validate template file
    if not os.path.isfile(template_path):
        raise FileNotFoundError(f"Template file does not exist: {template_path}")

    src_path = template_path
    template_file = os.path.basename(src_path)

    # Build paths
    new_filename = f"{output_filename_prefix}.xlsx"
    sharepoint_path = os.path.join(sharepoint_folder, new_filename)
    desktop_path = os.path.join(desktop_folder, new_filename)

    print("=" * 60)
    print("SHAREPOINT UPLOAD - STARTED")
    print("=" * 60)
    print(f"Template: {template_file}")
    print(f"Output: {new_filename}")
    print("=" * 60)

    # --- Step 1: Copy the template to Desktop as a working copy ---
    try:
        if os.path.exists(desktop_path):
            os.chmod(desktop_path, stat.S_IWRITE)
            os.remove(desktop_path)
        shutil.copyfile(src_path, desktop_path)
        os.chmod(desktop_path, stat.S_IWRITE)
        print("\n[OK] Step 1: Template copied to Desktop")
        print(f"   Location: {desktop_path}")
    except Exception as e:
        print(f"\n[Error] Step 1 failed: {e}")
        raise

    # --- Step 2: Open the working copy and write DataFrame ---
    if df is not None:
        try:
            app = xw.App(visible=excel_visible, add_book=False)
            app.display_alerts = False
            app.screen_updating = False

            try:
                wb = app.books.open(desktop_path)
                ws = wb.sheets[0]
                print("\n[OK] Step 2: Excel file opened")

                # Convert date/time columns to strings (xlwings can't handle them)
                df_write = df.copy()
                for col in df_write.columns:
                    if df_write[col].apply(lambda x: isinstance(x, (_dt.time, _dt.date)) and not isinstance(x, _dt.datetime)).any():
                        df_write[col] = df_write[col].astype(str)

                ws.range('A1').options(index=False).value = df_write
                print(f"   [OK] Wrote {df.shape[0]} rows x {df.shape[1]} cols")

                # Apply column formats
                if column_types and df.shape[0] > 0:
                    headers = [str(c) for c in df.columns.tolist()]
                    for col_name, fmt_type in column_types.items():
                        if col_name not in headers:
                            print(f"   [Warning] column_types: '{col_name}' not found — skipping")
                            continue
                        fmt_code = _EXCEL_FORMATS.get(fmt_type.lower())
                        if fmt_code is None:
                            print(f"   [Warning] column_types: unknown type '{fmt_type}' — skipping")
                            continue
                        col_idx = headers.index(col_name)
                        data_range = ws.range((2, col_idx + 1), (df.shape[0] + 1, col_idx + 1))
                        data_range.number_format = fmt_code
                        print(f"   [OK] column_types: '{col_name}' -> {fmt_type} ({fmt_code})")

                wb.save()
                wb.close()
                print("\n   [OK] All changes saved")

            finally:
                try:
                    app.quit()
                except Exception as quit_err:
                    print(f"   [Warning] Could not quit Excel (may already be closed): {quit_err}")

        except Exception as e:
            print(f"\n[Error] Step 2 failed: {e}")
            raise
    else:
        print("\n[Skip] Step 2: No data provided, file just gets copied")

    # --- Step 3: Copy to SharePoint ---
    try:
        shutil.copyfile(desktop_path, sharepoint_path)
        os.chmod(sharepoint_path, stat.S_IWRITE)
        print(f"\n[OK] Step 3: Copied to SharePoint")
        print(f"   Location: {sharepoint_path}")
    except Exception as e:
        print(f"\n[Error] Step 3 failed: {e}")
        print("   Check: permissions, checkout requirements, folder access")
        raise

    # Cleanup
    if not keep_desktop_copy:
        try:
            os.remove(desktop_path)
            print(f"\n[OK] Desktop copy removed")
        except Exception as e:
            print(f"\n[Warning] Could not remove desktop copy: {e}")

    print("\n" + "=" * 60)
    print(f"[Done] {sharepoint_path}")
    print("=" * 60 + "\n")

    return {
        'template_path': src_path,
        'sharepoint_path': sharepoint_path,
        'desktop_path': desktop_path,
        'filename': new_filename
    }
