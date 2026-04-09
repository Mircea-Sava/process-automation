"""Query any SAP table via RFC and print the result to a text file."""
from process_automation import sap_query

if __name__ == "__main__":
    df = sap_query(
        table="DD02L",                                                          # <-- CHANGE: any SAP table (DD02L, DD03L, KNA1, etc.)
        cols="TABNAME,TABCLASS,AS4USER,AS4DATE",                                # <-- CHANGE: comma-separated column names (empty = all)
        filters=[                                                               # <-- CHANGE: filter conditions
            "TABNAME LIKE 'Z%'",
        ],
        max_pages=0,                                                            # <-- CHANGE: 0 = all pages, limits rows fetched from SAP
    )
    # --- Output ----------------------------------------------------------------
    output_file  = r"C:\Temp\sap_query_result.txt"                              # <-- CHANGE: output file path

    result = df.to_string(index=False)
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(result)
    print(f"Result saved to: {output_file}  ({len(df)} rows)")
