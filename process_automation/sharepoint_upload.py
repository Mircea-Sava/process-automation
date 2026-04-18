# SharePoint File Uploader
# ========================
# Uploads files to SharePoint using a template-based approach. Supports both
# Excel (DataFrame written via xlwings) and CSV (DataFrame saved internally as
# CSV then uploaded via a workaround that avoids SharePoint checkout locks).
#
# Excel mode (default):
#   1. Copies the template .xlsx to Desktop as a working copy
#   2. Opens with xlwings and writes the DataFrame into the first sheet
#   3. Saves and copies the file to the SharePoint folder
#   4. Deletes the Desktop working copy (unless keep_desktop_copy=True)
#
# CSV mode:
#   1. Saves the DataFrame to a temp .csv locally
#   2. Copies template .xlsx to SharePoint as a placeholder
#   3. Renames temp .csv -> .xlsx (disguise)
#   4. Overwrites the placeholder with the disguised file
#   5. Renames the file on SharePoint from .xlsx -> .csv
#   6. Cleans up the temp file
#
# Why the CSV workaround: SharePoint triggers a checkout lock when an Office file
# (.xlsx, .docx) is created or overwritten via the mapped drive. By keeping .xlsx
# throughout and only renaming to .csv as the last step, SharePoint sees a plain
# file rename — no checkout lock.
#
# Usage (Excel):
#   from process_automation import save_excel_to_sharepoint
#
#   result = save_excel_to_sharepoint(df,
#       placeholder_template_path=r"Z:\path\to\TEMPLATE_DO_NOT_DELETE.xlsx",
#       sharepoint_folder=r"Z:\path\to\sharepoint_folder",
#       output_filename_prefix="MyReport_2026-02-26",
#   )
#
# Usage (CSV):
#   result = save_excel_to_sharepoint(df,
#       placeholder_template_path=r"Z:\path\to\TEMPLATE_DO_NOT_DELETE.xlsx",
#       sharepoint_folder=r"Z:\path\to\sharepoint_folder",
#       output_filename_prefix="MyReport_2026-02-26",
#       format="csv",
#   )

import os
import shutil
import stat
import datetime as _dt
from typing import Dict, Optional
import pandas as pd
import xlwings as xw

_TEMPLATE_PATTERN = "TEMPLATE_DO_NOT_DELETE"  # used by auto-detection if implemented later

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
    df: Optional[pd.DataFrame],
    placeholder_template_path: str,
    sharepoint_folder: str,
    output_filename_prefix: str,
    format: str = "xlsx",
    keep_desktop_copy: bool = False,
    excel_visible: bool = False,
    column_types: Optional[Dict[str, str]] = None,
    create_folders: bool = True,
) -> Dict[str, str]:

    if format not in ("xlsx", "csv"):
        raise ValueError("format must be 'xlsx' or 'csv'")

    if format == "csv":
        if df is None:
            raise ValueError("df is required when format='csv'")
        return _upload_csv_workaround(
            df=df,
            placeholder_template_path=placeholder_template_path,
            sharepoint_folder=sharepoint_folder,
            output_filename_prefix=output_filename_prefix,
            create_folders=create_folders,
            keep_desktop_copy=keep_desktop_copy,
        )

    return _upload_excel(
        df=df,
        placeholder_template_path=placeholder_template_path,
        sharepoint_folder=sharepoint_folder,
        output_filename_prefix=output_filename_prefix,
        keep_desktop_copy=keep_desktop_copy,
        excel_visible=excel_visible,
        column_types=column_types,
        create_folders=create_folders,
    )


def _upload_excel(
    df: Optional[pd.DataFrame],
    placeholder_template_path: str,
    sharepoint_folder: str,
    output_filename_prefix: str,
    keep_desktop_copy: bool,
    excel_visible: bool,
    column_types: Optional[Dict[str, str]],
    create_folders: bool,
) -> Dict[str, str]:

    desktop_folder = os.path.join(os.environ["USERPROFILE"], "Desktop")
    try:
        os.makedirs(sharepoint_folder, exist_ok=True)
    except OSError as e:
        print(f"   [Warning] Could not create SharePoint folder dynamically. Relying on existing path. ({e})")
    os.makedirs(desktop_folder, exist_ok=True)

    if not os.path.isfile(placeholder_template_path):
        raise FileNotFoundError(f"Template file does not exist: {placeholder_template_path}")

    src_path = placeholder_template_path
    template_file = os.path.basename(src_path)

    new_filename = f"{output_filename_prefix}.xlsx"
    sharepoint_path = os.path.join(sharepoint_folder, new_filename)
    desktop_path = os.path.join(desktop_folder, new_filename)

    print("=" * 60)
    print("SHAREPOINT UPLOAD (XLSX) - STARTED")
    print("=" * 60)
    print(f"Template: {template_file}")
    print(f"Output: {new_filename}")
    print("=" * 60)

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

    if df is not None:
        try:
            app = xw.App(visible=excel_visible, add_book=False)
            app.display_alerts = False
            app.screen_updating = False

            try:
                wb = app.books.open(desktop_path)
                ws = wb.sheets[0]
                print("\n[OK] Step 2: Excel file opened")

                df_write = df.copy()
                for col in df_write.columns:
                    if df_write[col].apply(lambda x: isinstance(x, (_dt.time, _dt.date)) and not isinstance(x, _dt.datetime)).any():
                        df_write[col] = df_write[col].astype(str)

                ws.range('A1').options(index=False).value = df_write
                print(f"   [OK] Wrote {df.shape[0]} rows x {df.shape[1]} cols")

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

    try:
        shutil.copyfile(desktop_path, sharepoint_path)
        os.chmod(sharepoint_path, stat.S_IWRITE)
        print(f"\n[OK] Step 3: Copied to SharePoint")
        print(f"   Location: {sharepoint_path}")
    except Exception as e:
        print(f"\n[Error] Step 3 failed: {e}")
        print("   Check: permissions, checkout requirements, folder access")
        raise

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
        'placeholder_template_path': src_path,
        'sharepoint_path': sharepoint_path,
        'desktop_path': desktop_path,
        'filename': new_filename
    }


def _upload_csv_workaround(
    df: pd.DataFrame,
    placeholder_template_path: str,
    sharepoint_folder: str,
    output_filename_prefix: str,
    create_folders: bool,
    keep_desktop_copy: bool = False,
) -> Dict[str, str]:

    if not os.path.isfile(placeholder_template_path):
        raise FileNotFoundError(f"Template file does not exist: {placeholder_template_path}")

    if create_folders:
        try:
            os.makedirs(sharepoint_folder, exist_ok=True)
        except OSError as e:
            print(f"   [Warning] Could not create SharePoint folder dynamically. ({e})")

    template_file = os.path.basename(placeholder_template_path)

    sp_placeholder_xlsx = os.path.join(sharepoint_folder, f"{output_filename_prefix}.xlsx")
    sp_final_csv = os.path.join(sharepoint_folder, f"{output_filename_prefix}.csv")

    desktop_folder = os.path.join(os.environ["USERPROFILE"], "Desktop")
    os.makedirs(desktop_folder, exist_ok=True)

    print("=" * 60)
    print("SHAREPOINT UPLOAD (CSV) - STARTED")
    print("=" * 60)
    print(f"Template : {template_file}")
    print(f"Output   : {output_filename_prefix}.csv")
    print(f"Rows     : {df.shape[0]:,} x {df.shape[1]}")
    print("=" * 60)

    temp_csv = os.path.join(desktop_folder, f"{output_filename_prefix}_temp.csv")

    try:
        df.to_csv(temp_csv, index=False)
        print(f"\n[OK] Step 1: DataFrame saved to Desktop as temp CSV")
        print(f"   Location: {temp_csv}")

        for leftover in (sp_placeholder_xlsx, sp_final_csv):
            if os.path.exists(leftover):
                os.remove(leftover)

        shutil.copyfile(placeholder_template_path, sp_placeholder_xlsx)
        print(f"\n[OK] Step 2: Template copied to SharePoint as .xlsx placeholder")
        print(f"   Location: {sp_placeholder_xlsx}")

    except Exception as e:
        print(f"\n[Error] Step 1-2 failed: {e}")
        raise

    local_xlsx = os.path.splitext(temp_csv)[0] + ".xlsx"

    try:
        if os.path.exists(local_xlsx):
            os.remove(local_xlsx)
        os.rename(temp_csv, local_xlsx)
        print(f"\n[OK] Step 3: Temp file renamed from .csv to .xlsx")
        print(f"   Location: {local_xlsx}")
    except Exception as e:
        print(f"\n[Error] Step 3 failed: Could not rename temp file to .xlsx: {e}")
        if not keep_desktop_copy:
            for f in (local_xlsx, temp_csv):
                try:
                    if os.path.exists(f):
                        os.remove(f)
                except Exception:
                    pass
        raise

    try:
        shutil.copyfile(local_xlsx, sp_placeholder_xlsx)
        print(f"\n[OK] Step 4: Local .xlsx pasted into SharePoint (overwrites placeholder)")
        print(f"   Source : {local_xlsx}")
        print(f"   Dest   : {sp_placeholder_xlsx}")
    except Exception as e:
        print(f"\n[Error] Step 4 failed: {e}")
        raise

    try:
        os.rename(sp_placeholder_xlsx, sp_final_csv)
        print(f"\n[OK] Step 5: File renamed to .csv in SharePoint")
        print(f"   Location: {sp_final_csv}")
    except Exception as e:
        print(f"\n[Error] Step 5 failed: Could not rename SharePoint file to .csv: {e}")
        raise

    if keep_desktop_copy:
        print(f"\n[OK] Step 6: Desktop copies kept")
        print(f"   CSV: {temp_csv}")
        print(f"   XLSX: {local_xlsx}")
    else:
        try:
            os.remove(local_xlsx)
        except Exception as e:
            print(f"\n[Warning] Could not remove Desktop .xlsx (non-blocking): {e}")
        try:
            if os.path.exists(temp_csv):
                os.remove(temp_csv)
        except Exception as e:
            print(f"[Warning] Could not remove Desktop .csv (non-blocking): {e}")
        print(f"\n[OK] Step 6: Desktop copies removed")

    print("\n" + "=" * 60)
    print("[Done] PROCESS COMPLETED")
    print("=" * 60)
    print(f"Final file location: {sp_final_csv}")
    print("=" * 60 + "\n")

    result = {
        "placeholder_template_path": placeholder_template_path,
        "sharepoint_path": sp_final_csv,
        "filename": f"{output_filename_prefix}.csv",
        "desktop_csv_path": temp_csv,
        "desktop_xlsx_path": local_xlsx,
    }
    return result