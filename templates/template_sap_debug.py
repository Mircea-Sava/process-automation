"""Debug a SAP GUI script — on failure, captures a screenshot + element map."""
from process_automation import sap_debug

if __name__ == "__main__":
    sap_debug(
        script_path=r"C:\path\to\your_template.py",                            # <-- CHANGE: path to the SAP template to debug
        output_dir=r"C:\Temp",                                                  # <-- CHANGE: where to save debug files
    )
