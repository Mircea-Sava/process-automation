# SAP Extract Engine
# ==================
# Orchestrates the full SAP extract workflow.
# Called by each transaction script (ZSUPVENG.py, ZMB51.py, ZSUPV.py, etc.) via run_extract().
#
# What run_extract() does (in order):
#   1. Pre-process (optional)  : copies a file from SharePoint, filters rows, extracts one column
#                                 into a _copy file that SAP can import as a selection filter
#   2. SAP connection          : opens SAP GUI and connects to PR1 (via sap_connection.py)
#   3. Transaction             : enters the transaction code and runs your sap_script(session) function
#   4. Save dialog             : fills in the file save dialog (path + filename) for txt/xlsx exports
#   5. Read into memory        : loads the exported file into a pandas DataFrame
#   6. Close SAP               : logs off and closes the SAP window
#   7. SharePoint upload       : uploads the DataFrame to SharePoint as xlsx (via sharepoint_upload.py)
#   8. Cleanup                 : deletes local file and/or preprocessed copy if requested
#   9. Returns the DataFrame   : so your script can continue working with the data
#
# Usage:
#   from process_automation import run_extract
#
#   df = run_extract(
#       sap_script,
#       transaction="ZSUPVENG",
#       template_folder=r"Z:\path\to\template_folder",
#       export_format="xlsx",
#       upload_to_sharepoint=True,
#       sharepoint_folder=r"Z:\00_DATABASE_DLDP\ADO_TEMPLATE\ZSUPVENG",
#       ...
#   )
#
# See templates/template_sap_pipeline.py for a full example with all parameters.

import os
import shutil
import time
from datetime import datetime
import pandas as pd
from .sap_connection import SAPManager
from .sharepoint_upload import save_excel_to_sharepoint


def build_filename(base, use_date=True, use_time=True):
    """Build a filename from base name with optional date and time suffixes."""
    parts = [base]
    now = datetime.now()
    if use_date:
        parts.append(now.strftime("%Y-%m-%d"))
    if use_time:
        parts.append(now.strftime("%H-%M-%S"))
    return "_".join(parts)


def run_extract(sap_script, transaction="", export_format="xlsx",
                template_folder="",
                sharepoint_folder="",
                sharepoint_filename="Default", download_dir=None,
                download_filename="",
                download_use_date=None, download_use_time=None,
                sharepoint_use_date="Default", sharepoint_use_time="Default",
                upload_to_sharepoint=True, delete_after_upload=False,
                preprocess_file=False, preprocess_source=None,
                preprocess_dest_dir=None,
                preprocess_column=None, preprocess_unique=True,
                preprocess_filters=None, preprocess_skiprows=None,
                delete_preprocess_copy=False,
                column_types=None):

    if download_dir is None:
        download_dir = os.path.join(os.environ['USERPROFILE'], "Downloads")

    # Build output filename
    dl_base = download_filename or transaction
    dl_date = download_use_date if download_use_date is not None else sharepoint_use_date
    dl_time = download_use_time if download_use_time is not None else sharepoint_use_time
    OUTPUT_NAME = build_filename(dl_base, dl_date, dl_time)
    EXTENSIONS = {"txt": ".txt", "xlsx": ".xlsx"}
    extract_name2 = f"{OUTPUT_NAME}{EXTENSIONS[export_format]}" if export_format in EXTENSIONS else None

    # --- Pre-process: copy and process a file from a previous extract ---
    preprocess_copy_path = None
    if preprocess_file:
        if not preprocess_source or not os.path.exists(preprocess_source):
            raise ValueError(f"preprocess_source file not found: {preprocess_source}")
        if not preprocess_dest_dir:
            raise ValueError("preprocess_dest_dir is required when preprocess_file=True")

        src_name, src_ext_full = os.path.splitext(os.path.basename(preprocess_source))
        dest_filename = f"{src_name}_copy{src_ext_full}"
        dest_path = os.path.join(preprocess_dest_dir, dest_filename)
        src_ext = os.path.splitext(preprocess_source)[1].lower()

        # Copy source file to destination
        shutil.copy2(preprocess_source, dest_path)
        print(f"Copied {preprocess_source} -> {dest_path}")

        if preprocess_column:
            if src_ext == ".txt":
                # TXT: extract a single column and overwrite the file
                if preprocess_skiprows is None:
                    with open(dest_path, 'r', encoding='latin-1') as f:
                        for i, line in enumerate(f):
                            if preprocess_column.strip() in [c.strip() for c in line.split('\t')]:
                                preprocess_skiprows = i
                                break
                        else:
                            raise ValueError(f"Column '{preprocess_column}' not found in {dest_path}")
                    print(f"Auto-detected header at row {preprocess_skiprows}")

                df_reread = pd.read_csv(dest_path, sep='\t', encoding='latin-1',
                                        skiprows=preprocess_skiprows,
                                        on_bad_lines='skip', engine='python')
                df_reread.columns = df_reread.columns.str.strip()
                print("Columns:", df_reread.columns.tolist())
                if preprocess_filters:
                    for col, vals in preprocess_filters.items():
                        df_reread[col] = df_reread[col].astype(str).str.strip()
                        df_reread = df_reread[df_reread[col].isin([str(v) for v in vals])]
                        print(f"Filtered '{col}' to {vals}: {len(df_reread)} rows remaining")
                values = df_reread[preprocess_column].dropna().astype(int).astype(str)
                if preprocess_unique:
                    values = values.unique()
                with open(dest_path, 'w') as f:
                    f.write('\n'.join(values))
                print(f"Pre-processed txt: kept {'unique ' if preprocess_unique else ''}values from '{preprocess_column}'")

            elif src_ext == ".xlsx":
                # XLSX: extract a single column and save as CSV
                df_reread = pd.read_excel(dest_path)
                df_reread.columns = df_reread.columns.str.strip()
                print("Columns:", df_reread.columns.tolist())
                if preprocess_filters:
                    for col, vals in preprocess_filters.items():
                        df_reread[col] = df_reread[col].astype(str).str.strip()
                        df_reread = df_reread[df_reread[col].isin([str(v) for v in vals])]
                        print(f"Filtered '{col}' to {vals}: {len(df_reread)} rows remaining")
                values = df_reread[preprocess_column].dropna().astype(int).astype(str)
                if preprocess_unique:
                    values = values.unique()
                csv_path = dest_path.replace(".xlsx", ".csv")
                with open(csv_path, 'w') as f:
                    f.write('\n'.join(values))
                os.remove(dest_path)
                dest_path = csv_path
                print(f"Pre-processed xlsx -> csv: kept {'unique ' if preprocess_unique else ''}values from '{preprocess_column}'")

        preprocess_copy_path = dest_path
        print(f"Pre-processed file ready: {dest_path}")

    # --- SAP Connection ---
    sap = SAPManager()
    session = sap.get_session()
    print("SAP session obtained successfully.")

    file_path = None
    df_copied = None
    try:
        # Enter the transaction
        session.findById("wnd[0]/tbar[0]/okcd").text = transaction
        session.findById("wnd[0]").sendVKey(0)

        # Run the transaction-specific SAP script
        sap_script(session)

        # --- Save dialog (sap_script should have triggered the export menu already) ---
        if export_format == "txt":
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = download_dir
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = extract_name2
            session.findById("wnd[1]").sendVKey(0)

        elif export_format == "xlsx":
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = download_dir
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = extract_name2
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 4
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Close the specific Excel process that SAP opened for the export.
        if export_format == "xlsx":
            file_target = os.path.join(download_dir, extract_name2)
            # Wait up to 30s for the file to appear and stabilize
            for _ in range(30):
                if os.path.exists(file_target):
                    try:
                        size1 = os.path.getsize(file_target)
                        time.sleep(1)
                        size2 = os.path.getsize(file_target)
                        if size2 == size1 and size2 > 0:
                            break
                    except OSError:
                        pass
                time.sleep(1)

            # Find and kill only the Excel process that has this specific file open
            try:
                import psutil
                target_path = os.path.normcase(os.path.abspath(file_target))
                for proc in psutil.process_iter(['name', 'pid']):
                    if proc.info['name'] and proc.info['name'].lower() == 'excel.exe':
                        try:
                            for f in proc.open_files():
                                if os.path.normcase(f.path) == target_path:
                                    proc.kill()
                                    proc.wait(timeout=5)
                                    print(f"Killed Excel (PID {proc.pid}) holding {extract_name2}")
                                    break
                        except (psutil.NoSuchProcess, psutil.AccessDenied):
                            continue
            except Exception as e:
                print(f"[Warning] Could not close SAP Excel: {e}")

        # Read result into DataFrame
        if export_format == "txt":
            file_path = os.path.join(download_dir, extract_name2)
            df_copied = pd.read_csv(file_path, sep='\t', encoding='latin-1', on_bad_lines='skip', engine='python')
        elif export_format == "xlsx":
            file_path = os.path.join(download_dir, extract_name2)
            df_copied = pd.read_excel(file_path)
            # Fix: convert Timedelta columns to string (SAP duration fields cause this)
            for col in df_copied.select_dtypes(include=['timedelta', 'timedelta64']).columns:
                df_copied[col] = df_copied[col].astype(str)
        elif export_format == "clipboard":
            df_copied = pd.read_clipboard(sep='\t', on_bad_lines='skip', engine='python')
    finally:
        # --- Close SAP (always, even if an error occurred) ---
        try:
            sap.close_connection(session)
        except Exception as e:
            print(f"[Warning] Could not close SAP session: {e}")

    # --- SharePoint Upload ---
    if df_copied is None:
        print("[Error] No data was extracted — skipping upload.")
        return None

    if upload_to_sharepoint:
        # Resolve "Default" sharepoint settings to reuse download block values
        if sharepoint_filename == "Default":
            sharepoint_filename = download_filename
        if sharepoint_use_date == "Default":
            sharepoint_use_date = download_use_date if download_use_date is not None else True
        if sharepoint_use_time == "Default":
            sharepoint_use_time = download_use_time if download_use_time is not None else True

        sp_filename = sharepoint_filename or transaction
        OUTPUT_PREFIX = build_filename(sp_filename, sharepoint_use_date, sharepoint_use_time)

        result = save_excel_to_sharepoint(
            df_copied,
            template_folder=template_folder,
            sharepoint_folder=sharepoint_folder,
            output_filename_prefix=OUTPUT_PREFIX,
            column_types=column_types
        )
        print(f"\n[Success] Process completed successfully!")
        print(f"Final file location: {result['sharepoint_path']}")

        # Clean up local file
        if delete_after_upload and file_path and os.path.exists(file_path):
            max_retries = 3
            for i in range(max_retries):
                try:
                    os.remove(file_path)
                    print(f"Deleted temporary file: {file_path}")
                    break
                except PermissionError as e:
                    if i < max_retries - 1:
                        print(f"   [Warning] File locked, retrying in 2 seconds... ({e})")
                        time.sleep(2)
                    else:
                        print(f"   [Warning] Could not delete temporary file, it might be locked: {file_path}")
    else:
        print(f"\n[Success] Process completed successfully!")
        if file_path:
            print(f"File saved locally: {file_path}")

    # --- Clean up preprocessed copy ---
    if delete_preprocess_copy and preprocess_copy_path and os.path.exists(preprocess_copy_path):
        os.remove(preprocess_copy_path)
        print(f"Deleted preprocessed copy: {preprocess_copy_path}")

    return df_copied
