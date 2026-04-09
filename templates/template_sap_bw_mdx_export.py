"""Export a SAP BW query (BEx / InfoProvider) to CSV/XLSX via MDX + OLAP BAPIs."""
from process_automation import run_bw_mdx_extract

if __name__ == "__main__":
    run_bw_mdx_extract({
        # --- Option A: structured (no MDX knowledge required) ----------------
        "query":   "ZMY_QUERY",                                  # <-- CHANGE: released BEx query (technical name)
        "rows":    ["[0MATERIAL].[LEVEL01].MEMBERS"],            # <-- CHANGE: row-axis member set(s)
        "columns": ["[Measures].[ZSALES]", "[Measures].[ZQTY]"], # <-- CHANGE: column-axis member set(s)
        "where":   "([0CALMONTH].[202603])",                     # <-- CHANGE: optional slicer
        "variables": {                                            # <-- CHANGE: SAP query variables
            "ZVAR_PLANT": "0005",
        },

        # --- Option B: raw MDX (overrides the structured fields) -------------
        # "mdx": "SELECT {[Measures].[ZSALES]} ON COLUMNS, "
        #        "{[0MATERIAL].[LEVEL01].MEMBERS} ON ROWS "
        #        "FROM [QUERYCUBE/ZMY_QUERY] "
        #        "SAP VARIABLES [ZVAR_PLANT] INCLUDING '0005'",

        # --- Where to save ---------------------------------------------------
        "output_dir":  r"Z:\00_DATABASE_DLDP\ADO_TEMPLATE",      # <-- CHANGE: output directory
        "output_name": "ZMY_QUERY",                              # <-- CHANGE: filename, or None to use query name
        "extension":   "xlsx",                                   # <-- CHANGE: "csv" or "xlsx"
        "add_date":    True,                                     # <-- CHANGE: append _YYYYMMDD
        "add_time":    True,                                     # <-- CHANGE: append _HHMMSS
    })
