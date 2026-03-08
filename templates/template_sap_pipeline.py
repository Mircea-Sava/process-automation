import os
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).resolve().parent / "functions"))
from sap_extract import run_extract


def sap_script(session):
    """Paste your full SAP recording here.

    Include everything from after the transaction opens, through running
    the report, up to and including the export menu click.
    Stop before the file save dialog — the engine handles that automatically.
    """

    # Replace everything below with your own SAP recording
    session.findById("wnd[0]/usr/ctxtS_PERNR-LOW").text = "29365"
    session.findById("wnd[0]/usr/ctxtS_PERNR-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtS_PERNR-LOW").caretPosition = 8
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()


if __name__ == "__main__":
    run_extract(sap_script,

        # --- Preprocess: prepare file to import (optional) -------------------
        #
        # Set to True when your SAP transaction needs "Import from text file".
        # Copies a file from SharePoint, extracts one column, and converts it
        # to a SAP-friendly format (.txt or .csv).
        #
        # If source is .txt  -> extracts column and saves as .txt copy
        # If source is .xlsx -> extracts column and saves as .csv copy

        preprocess_file=False,                                                  # <-- CHANGE: True = enable preprocessing
        preprocess_source=None,                                                 # <-- CHANGE: full path to source file (e.g. os.path.join(r"Z:\...\ZSUPV", f"ZSUPV_{datetime.now().strftime('%Y-%m-%d')}.xlsx"))
        preprocess_dest_dir=None,                                               # <-- CHANGE: directory for processed copy (e.g. os.path.join(os.environ["USERPROFILE"], "Downloads"))
        preprocess_filters=None,                                                # <-- CHANGE: filter rows before extracting column (e.g. {"Status": ["Active", "Pending"]})
        preprocess_column=None,                                                 # <-- CHANGE: column name to extract (e.g. "Pers.No.")
        preprocess_unique=False,                                                # <-- CHANGE: True = remove duplicates

        # --- Transaction settings --------------------------------------------

        transaction="zsupv",                                                    # <-- CHANGE: your transaction code

        # --- Export format ---------------------------------------------------
        #
        # Tell the engine what file format your export produces.
        # The export menu clicks should be in sap_script() above.

        export_format="xlsx",                                                   # <-- CHANGE: "txt", "xlsx", or "clipboard"

        # --- Local download settings -----------------------------------------

        download_dir=os.path.join(os.environ["USERPROFILE"], "Downloads"),      # <-- CHANGE: local download directory
        download_filename="zsupv",                                              # <-- CHANGE: base filename (leave "" to use transaction name)
        download_use_date=True,                                                 # <-- CHANGE: True = append date to filename
        download_use_time=True,                                                 # <-- CHANGE: True = append time to avoid filename conflicts

        # --- SharePoint settings ---------------------------------------------

        upload_to_sharepoint=True,                                             # <-- CHANGE: True = upload to SharePoint, False = keep local only
        template_folder=r"Z:\path\to\template_folder",                         # <-- CHANGE: folder containing your template .xlsx file
        sharepoint_folder=r"Z:\path\to\your_sharepoint_folder",                # <-- CHANGE: SharePoint destination folder
        sharepoint_filename="Default",                                         # <-- CHANGE: "Default" = same as download_filename, or set a custom name
        sharepoint_use_date="Default",                                         # <-- CHANGE: "Default" = same as download_use_date, or True/False
        sharepoint_use_time="Default",                                         # <-- CHANGE: "Default" = same as download_use_time, or True/False

        # --- Column type overrides (optional) --------------------------------

        column_types=None,                                                     # <-- CHANGE: e.g. {"Amount": "currency", "Date": "date"}
                                                                               #     Types: "number", "currency", "date", "time", "percentage", "fraction", "text", "general"

        # --- Cleanup ---------------------------------------------------------

        delete_after_upload=False,                                             # <-- CHANGE: True = delete local file after upload
        delete_preprocess_copy=False,                                          # <-- CHANGE: True = delete preprocessed copy after extract
    )
