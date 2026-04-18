# Process Automation

**Most powerful feature: the RFC tool lets you query practically any raw SAP table directly** — even if your company has restricted access to certain transactions, you can go under the hood. No screen access needed, just table name and filters.

Beyond RFC, automate SAP GUI extractions, Databricks queries, web scraping, SharePoint uploads and Excel macro execution with simple Python templates.

## Install

```
pip install process-automation
```

## Templates

Copy a template, fill in your settings, and run it.

| Template | What it does |
|----------|-------------|
| `template_sap_pipeline.py` | Record a SAP transaction, export data, upload to SharePoint |
| `template_sap_rfc_export.py` | Export SAP tables directly via RFC to CSV/XLSX |
| `template_sap_rfc_schema.py` | Query any SAP table (DD02L, DD03L, etc.) and print results |

| `template_sap_debug.py` | Debug a SAP script — captures screenshot + element map on failure |
| `template_sharepoint_upload.py` | Upload a DataFrame to SharePoint as formatted Excel |
| `template_databricks_pipeline.py` | Query Databricks and save locally or upload to SharePoint |
| `template_excel_macro.py` | Run VBA macros in an Excel workbook |

## Quick Examples

**SAP GUI pipeline** — paste your SAP recording, run the script:

```python
from process_automation import run_extract

df = run_extract(sap_script,
    export_format="xlsx",
    template_path=r"Z:\path\to\TEMPLATE.xlsx",
    sharepoint_folder=r"Z:\path\to\output",
)
```

**SAP RFC export** — no GUI needed, just table name and filters:

```python
from process_automation import run_rfc_extract

run_rfc_extract({
    "table": "AFRU",
    "cols": "PERNR,AUFNR,VORNR,BUDAT",
    "filters": ["BUDAT >= '20260301'"],
    "output_dir": r"Z:\path\to\output",
    "extension": "xlsx",
})
```

**SAP debug** — point it at a template, get diagnostics on failure:

```python
from process_automation import sap_debug

sap_debug(script_path=r"C:\path\to\your_template.py", output_dir=r"C:\Temp")
```

## Setup

- [SAP_RECORDING_GUIDE.md](SAP_RECORDING_GUIDE.md) — SAP GUI recording setup and RFC credential management
- [DATABRICKS_SETUP.md](DATABRICKS_SETUP.md) — Databricks connection and authentication
- [WEBSCRAPING_GUIDE.md](WEBSCRAPING_GUIDE.md) — Web scraping setup with Playwright

## Credits

Originally started as an internal team effort; packaged and open-sourced by contributors from the same group.

## License

MIT
