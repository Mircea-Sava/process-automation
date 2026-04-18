"""Export a SAP table to CSV or XLSX via RFC (no SAP GUI needed)."""
from process_automation import run_rfc_extract

if __name__ == "__main__":
    run_rfc_extract({
        # --- OPTIONAL: chained two-step export --------------------------------
        # "chain": {
        #     "table": "AUFK", "cols": "AUFNR,OBJNR,WERKS,ERDAT",
        #     "filters": ["ERDAT >= '20260319'", "(WERKS = '0005' OR WERKS = '0212')"],
        #     "source_column": "Object number", "target_field": "OBJNR",
        #     "save_file": True, "save_dir": r"C:\Temp", "save_name": "AUFK_chain", "save_extension": "csv",
        # },

        # --- What to export ---------------------------------------------------
        "table":      "AFRU",                                                   # <-- CHANGE: SAP table name
        "cols":       "PERNR,AUFNR,VORNR,ARBID,LTXA1,ISMNW,BUDAT,ERSDA,AUERU", # <-- CHANGE: comma-separated column names
        "filters":    [                                                          # <-- CHANGE: filter conditions
            "BUDAT >= '20260302'",
            "BUDAT <= '20260323'",
        ],

        # --- Where to save ----------------------------------------------------
        "output_dir":  r"Z:\00_DATABASE_DLDP\ADO_TEMPLATE",                     # <-- CHANGE: output directory
        "output_name": "ZSUPV",                                                 # <-- CHANGE: custom filename, or None to use table name
        "extension":   "xlsx",                                                  # <-- CHANGE: "csv" or "xlsx"
        "add_date":    True,                                                    # <-- CHANGE: append _YYYYMMDD
        "add_time":    True,                                                    # <-- CHANGE: append _HHMMSS

        # --- OPTIONAL ---------------------------------------------------------
        # "use_fieldnames": True,                                              # <-- CHANGE: first DATA row contains field names (for BAPI_MDDATASET_* style queries)
        # "file_filter": {"path": r"Z:\my_file.xlsx", "column": "OBJNR", "field": "OBJNR"},
    })
