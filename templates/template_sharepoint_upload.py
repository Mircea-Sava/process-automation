import pandas as pd
from process_automation import save_excel_to_sharepoint


# --- Step 1: Load your data -------------------------------------------------

df = pd.read_excel(r"Z:\path\to\your_file.xlsx")                        # <-- CHANGE: path to your file


# --- Step 2: Upload to SharePoint -------------------------------------------

result = save_excel_to_sharepoint(df,

    # --- Destination ---------------------------------------------------------
    placeholder_template_path=r"Z:\path\to\template.xlsx",    # <-- CHANGE: empty .xlsx file (TEMPLATE_DO_NOT_DELETE.xlsx)
    sharepoint_folder=r"Z:\path\to\your_sharepoint_folder",              # <-- CHANGE: SharePoint destination folder
    output_filename_prefix="MyReport_2026-01-01",                        # <-- CHANGE: output filename (without extension)

    # --- Format: "xlsx" (default) or "csv" -----------------------------------
    # xlsx: writes DataFrame to Excel template via xlwings
    # csv:  saves DataFrame to Desktop as temp CSV and uploads via workaround
    format="xlsx",                                                      # <-- CHANGE: "xlsx" or "csv"

    # --- Column type overrides (optional, xlsx mode only) -------------------
    column_types=None,                                                   # <-- CHANGE: e.g. {"Amount": "currency", "Date": "date"}
                                                                          #     Types: "number", "currency", "date", "time", "percentage", "fraction", "text", "general"

    # --- Options -------------------------------------------------------------
    keep_desktop_copy=False,                                             # <-- CHANGE: True = keep working copy on Desktop
    excel_visible=False,                                                 # <-- CHANGE: True = show Excel while writing (xlsx mode only)
)

print(f"Uploaded to: {result['sharepoint_path']}")