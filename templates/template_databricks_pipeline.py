import os
from process_automation import run_databricks_extract


if __name__ == "__main__":
    run_databricks_extract(

        # --- Databricks source ----------------------------------------------
        #
        # .env keys used by default (process_automation/.env):
        # DATABRICKS_HOST (or DATABRICKS_SERVER_HOSTNAME)
        # DATABRICKS_TOKEN
        # DATABRICKS_HTTP_PATH
        #
        # Optional fallback if HTTP path is not set:
        # DATABRICKS_ORG_ID + DATABRICKS_CLUSTER_ID
        env_path=None,                                                        # <-- CHANGE: custom .env path, or None to use process_automation/.env

        table_path="team_aftermarket_mro.mro_all.vwf_sap_merged_aufk_afko_afvc_afvv",  # <-- CHANGE: source table used when query=None
        query=None,                                                           # <-- CHANGE: full SQL query string; when set, select_columns/filters/order_by/limit are ignored
        # query supports full SQL (CTEs, JOINs, GROUP BY, window functions, UNION, subqueries, etc.).
        # Keep it as a single SELECT statement for this workflow.
        # query="""
        # SELECT
        #     AUFNR,
        #     WERKS,
        #     GLTRP,
        #     NETWR
        # FROM team_aftermarket_mro.mro_all.vwf_sap_merged_aufk_afko_afvc_afvv
        # WHERE WERKS = '1000'
        #   AND GLTRP >= DATE '2026-01-01'
        # ORDER BY GLTRP DESC
        # LIMIT 50000
        # """,

        # --- Query shaping ---------------------------------------------------
        # Used only when query=None (auto-builds SQL as: SELECT ... FROM table_path WHERE ... ORDER BY ... LIMIT ...)
        select_columns=None,                                                  # <-- CHANGE: columns to return; None = all columns (*). Example: ["AUFNR", "WERKS", "GLTRP"]
        filters=None,                                                         # <-- CHANGE: WHERE conditions joined with AND. Example: ["WERKS = '1000'", "AUART IN ('PM01', 'PM02')"]
        order_by=None,                                                        # <-- CHANGE: sort order (string or list). Example: ["GLTRP DESC", "AUFNR ASC"]
        limit=None,                                                           # <-- CHANGE: max rows returned. Example: 100000. Leave None for no limit

        # --- Local output ----------------------------------------------------
        save_local=True,                                                      # <-- CHANGE: True = save file locally
        output_format="xlsx",                                                 # <-- CHANGE: "xlsx", "csv", or "parquet"
        download_dir=os.path.join(os.environ["USERPROFILE"], "Downloads", "Extract"),    # <-- CHANGE: final local output folder
        download_filename="Test",                                             # <-- CHANGE: base filename
        download_use_date=True,                                               # <-- CHANGE: True = append date
        download_use_time=False,                                              # <-- CHANGE: True = append time

        # --- SharePoint output ----------------------------------------------
        upload_to_sharepoint=False,                                           # <-- CHANGE: True = also upload to SharePoint
        template_path=r"Z:\path\to\TEMPLATE_DO_NOT_DELETE.xlsx",              # <-- CHANGE: full path to your template .xlsx file
        sharepoint_folder=r"Z:\path\to\your_sharepoint_folder\MRO_EXTRACTS",  # <-- CHANGE: final SharePoint destination folder
        sharepoint_filename="Default",                                        # <-- CHANGE: "Default" = same as download_filename
        sharepoint_use_date="Default",                                        # <-- CHANGE: "Default" = same as download_use_date, or True/False
        sharepoint_use_time="Default",                                        # <-- CHANGE: "Default" = same as download_use_time, or True/False

        # --- SharePoint Excel formatting (optional) -------------------------
        column_types=None,                                                    # <-- CHANGE: e.g. {"NETWR": "currency", "GLTRP": "date"}. Types: "number", "currency", "date", "time", "percentage", "fraction", "text", "general".
    )
