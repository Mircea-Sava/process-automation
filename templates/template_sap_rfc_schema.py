"""Retrieve the schema (field list) of a SAP table via RFC."""
from process_automation import save_table_schema

if __name__ == "__main__":
    save_table_schema(
        table="ZSUPV",                                                          # <-- CHANGE: SAP table name
        output_dir=r"C:\Temp",                                                  # <-- CHANGE: output directory
        extension="csv",                                                        # <-- CHANGE: "csv" or "xlsx"
        field_names="both",                                                     # <-- CHANGE: "technical" | "label" | "both"
        show_example_row=True,                                                  # <-- CHANGE: True = fetch one example row
    )
