# SAP Automation

## What is this?

We regularly need to pull data from SAP, format it as Excel, and upload it to SharePoint. Doing this manually is slow and repetitive. These scripts automate the entire process so you can run a single Python file and have the data extracted, saved, and uploaded automatically.

## Who is this for?

Anyone who needs to extract SAP reports on a recurring basis. You do not need to be a programmer. The templates are designed so you only need to:

1. Record your SAP steps (SAP has a built-in script recorder)
2. Paste the recording into the template
3. Fill in a few settings (transaction code, file paths)
4. Run the script

## How is this different from other tools?

Existing Python packages for SAP automation or SharePoint uploads each cover only one piece of the workflow, and even within that piece they typically lack features like parallel execution safety, automatic popup handling, or template-based formatting. To get the full pipeline you'd have to combine multiple packages and write your own glue code.

This package is the full pipeline in a single install:

- **SAP extraction** — connects to SAP GUI via COM scripting, runs your recorded transaction, and exports the data
- **Data preprocessing** — reads the export into pandas, with optional column extraction, filtering, and deduplication
- **SharePoint upload** — copies a checked-in Excel template, writes the formatted data into it, and saves it to SharePoint

Other differences:

| | This package | Typical alternatives |
|---|---|---|
| **Scope** | End-to-end: SAP to SharePoint | SAP connection only, or Excel only |
| **Setup for users** | Paste a VBS recording, set a few variables | Learn a framework (Robot Framework) or write glue code |
| **Parallel execution** | Each script gets its own SAP session and Excel process | Usually single-session, no isolation |
| **Output formatting** | Template-based Excel with column formatting | Raw CSV or unformatted export |

## How it works

The scripts follow this workflow:

1. Open SAP and connect to PR1
2. Run your recorded SAP transaction (fill in fields, execute report, click export)
3. Save the exported file locally
4. Read the file into memory
5. Copy a checked-in Excel template from SharePoint, write the data into it, and save it back to SharePoint
6. Clean up temporary files

The "checked-in file" is an `.xlsx` template that lives in a SharePoint folder. You specify the path to this template in your script. The scripts copy it, fill it with your data, and save the result as a new file in your destination folder. This ensures every upload has consistent formatting.

## Folder Structure

```
SAP_Automation/
├── functions/     Core modules that do the actual work (you don't edit these)
├── templates/     Starter scripts you copy and customize for each SAP transaction
└── tools/         Utilities like the transaction checker
```

## Templates

### SAP Pipeline (template_sap_pipeline.py)

The main template. Handles the full workflow: connect to SAP, run your transaction, export data, upload to SharePoint.

```python
df = run_extract(sap_script,
    transaction="ZSUPVENG",
    export_format="xlsx",
    template_path=r"Z:\path\to\TEMPLATE_DO_NOT_DELETE.xlsx",
    sharepoint_folder=r"Z:\path\to\your_sharepoint_folder",
    column_names={"OldName": "NewName"}, # optional
)
```

**How to use:**
1. Copy this template and rename it (e.g. `ZSUPVENG.py`)
2. Paste your SAP recording into the `sap_script()` function
3. Set `transaction` to your transaction code
4. Set `template_path` to the full path of your checked-in `.xlsx` template file
5. Set `sharepoint_folder` to where you want the output file saved
6. (Optional) Set `column_names` to rename columns before uploading
7. Run the script

### SharePoint Upload (template_sharepoint_upload.py)

Uploads any data to SharePoint as a formatted Excel file. No SAP needed. Useful when you already have data in a file or DataFrame and just want to push it to SharePoint.

```python
result = save_excel_to_sharepoint(df,
    template_path=r"Z:\path\to\TEMPLATE_DO_NOT_DELETE.xlsx",
    sharepoint_folder=r"Z:\path\to\your_sharepoint_folder",
    output_filename_prefix="MyReport_2026-01-01",
)
```

### Databricks Pipeline (template_databricks_pipeline.py)

Pulls data from Databricks, applies optional filters, then saves locally and/or uploads to SharePoint using the same template-based upload flow.

```python
df = run_databricks_extract(
    table_path="catalog.schema.table_name",
    select_columns=["col1", "col2"],
    filters=["status = 'ACTIVE'"],
    download_dir=r"C:\Users\you\Downloads\MRO_EXTRACTS",
    upload_to_sharepoint=True,
    template_path=r"Z:\path\to\TEMPLATE_DO_NOT_DELETE.xlsx",
    sharepoint_folder=r"Z:\path\to\your_sharepoint_folder\MRO_EXTRACTS",
)
```

### Excel Macros (template_excel_macro.py)

Opens an Excel workbook and runs VBA macros by name.

```python
run_excel_macro(
    file_path=r"Z:\path\to\your_workbook.xlsm",
    macro_name="MyMacro",
    module_name="Module1"
)
```

## Core Modules (functions/)

These are the building blocks. You don't need to edit these files.

| File | What it does |
|------|-------------|
| `sap_extract.py` | The main engine. Connects to SAP, runs your recording, handles the save dialog, reads the file, uploads to SharePoint |
| `sap_connection.py` | Opens SAP GUI, connects to PR1, handles popups like "Multiple Logon" |
| `sharepoint_upload.py` | Copies the checked-in template, writes your data into it, saves it to SharePoint |
| `excel_macro.py` | Opens an Excel file and runs a VBA macro |
| `databricks_extract.py` | Queries Databricks tables/SQL and outputs to local and/or SharePoint |

## Export Formats

Your SAP recording should include the export menu clicks (right-click, select format, etc.). The `export_format` setting just tells the engine what file type to expect so it can handle the save dialog correctly.

| Format | What the engine does |
|--------|---------------------|
| `"txt"` | Fills in the save dialog, reads the tab-delimited text file |
| `"xlsx"` | Fills in the save dialog, waits for the file, closes the Excel window SAP opens |
| `"clipboard"` | No save dialog needed, reads directly from clipboard |

## Running Multiple Scripts at Once

These scripts are safe to run in parallel (e.g. a sequencer running multiple SAP transactions at the same time):

- **SAP Logon won't block the sequencer** — SAP Logon is launched independently so the sequencer doesn't wait for it to close before moving on to the next task
- **Excel crashes won't hide the real error** — if Excel stops responding mid-task, the script reports the actual problem instead of a confusing secondary error
- **Only your Excel gets closed** — when a script finishes, it only closes the Excel window it opened. Other scripts' Excel windows are left alone
- **Excel gets a chance to close cleanly** — the script asks Excel to close nicely first and waits 5 seconds. If it doesn't respond, then it force-closes it

## Tools

**Transaction Checker** (`tools/check_transactions.py`): Scans all your scripts for SAP transaction codes, then connects to SAP and verifies each one opens successfully. Run it with `check_transactions.bat`.
