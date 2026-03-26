# Web Scrape Engine
# =================
# Automates the full Playwright web scraping workflow.
# Called by each scraping script via run_web_scrape().
#
# What run_web_scrape() does (in order):
#   1. Browser launch     : opens Edge via Playwright (channel="msedge", no driver needed)
#   2. Navigation          : goes to your URL and waits for the page to load
#   3. Scraping            : runs your scrape_fn(page, download_dir) — your recorded clicks
#   4. Browser close       : shuts down the browser
#   5. Read into memory    : loads the downloaded file into a pandas DataFrame
#   6. Transform           : applies transform_fn, column_names, column_order (if provided)
#   7. SharePoint upload   : uploads the DataFrame to SharePoint as xlsx (via sharepoint_upload.py)
#   8. Cleanup             : deletes local file if requested
#   9. Returns DataFrame   : so your script can continue working with the data
#
# How to create a new scraping script:
#   1. Run:  python -m playwright codegen --channel msedge <YOUR_URL>
#   2. Click through the flow you want to automate in the browser that opens
#   3. Copy the generated code into a scrape_fn(page, download_dir) function
#   4. Use run_web_scrape() to handle everything else
#
# See templates/template_web_scrape.py for a full example with all parameters.

import os
from io import StringIO
import pandas as pd
from playwright.sync_api import sync_playwright
from .sap_extract import build_filename
from .sharepoint_upload import save_excel_to_sharepoint


def _read_file(file_path, read_format=None, read_kwargs=None, html_table_index=0):
    """Read a downloaded file into a DataFrame.

    Auto-detects format from extension when read_format is None.
    """
    if read_kwargs is None:
        read_kwargs = {}

    ext = os.path.splitext(file_path)[1].lower()

    # Resolve format
    fmt = read_format
    if fmt is None:
        fmt = {
            ".csv": "csv",
            ".xlsx": "excel",
            ".xls": "html",      # .xls from web apps is usually HTML
            ".tsv": "csv",
        }.get(ext)

    if fmt is None:
        raise ValueError(
            f"Cannot auto-detect format for '{ext}'. "
            f"Set read_format to 'csv', 'excel', or 'html'."
        )

    if fmt == "csv":
        defaults = {"encoding": "utf-8-sig"}
        if ext == ".tsv":
            defaults["sep"] = "\t"
        return pd.read_csv(file_path, **{**defaults, **read_kwargs})

    if fmt == "excel":
        return pd.read_excel(file_path, **read_kwargs)

    if fmt == "html":
        with open(file_path, encoding="utf-8") as f:
            html = f.read()
        tables = pd.read_html(StringIO(html), **read_kwargs)
        if not tables:
            raise ValueError(f"No HTML tables found in {file_path}")
        idx = min(html_table_index, len(tables) - 1)
        return tables[idx]

    raise ValueError(f"Unknown read_format: '{fmt}'. Use 'csv', 'excel', or 'html'.")


def run_web_scrape(
    scrape_fn,
    url,

    # --- Browser settings -----------------------------------------------
    headless=False,
    wait_until="networkidle",

    # --- Download settings ----------------------------------------------
    download_dir=None,

    # --- Data reading (when scrape_fn returns a file path) --------------
    read_format=None,
    read_kwargs=None,
    html_table_index=0,

    # --- Post-processing ------------------------------------------------
    transform_fn=None,
    column_names=None,
    column_order=None,

    # --- SharePoint upload ----------------------------------------------
    upload_to_sharepoint=False,
    template_path="",
    sharepoint_folder="",
    output_filename_prefix="",
    output_use_date=True,
    output_use_time=False,
    column_types=None,

    # --- Cleanup --------------------------------------------------------
    delete_after_upload=False,
):
    """Run a Playwright web scrape pipeline.

    Parameters
    ----------
    scrape_fn : callable
        Your scraping function.  Signature: ``scrape_fn(page, download_dir) -> str``
        *page* is a Playwright Page object, *download_dir* is the local folder
        for saving downloads.  Must return the **full path** of the downloaded file.

        Typical pattern inside scrape_fn::

            with page.expect_download(timeout=60000) as download_info:
                page.get_by_role("menuitem", name="Export to CSV", exact=True).click()
            download = download_info.value
            path = os.path.join(download_dir, download.suggested_filename)
            download.save_as(path)
            return path

    url : str
        The starting URL to navigate to.

    headless : bool
        False (default) shows the browser window — useful for debugging.

    wait_until : str
        Playwright page load strategy.  Default ``"networkidle"``.

    download_dir : str or None
        Local folder for saving downloads.  Defaults to ``~/Downloads``.

    read_format : str or None
        How to read the downloaded file into a DataFrame.
        ``None`` = auto-detect from file extension (.csv, .xlsx, .xls, .tsv).
        Explicit options: ``"csv"``, ``"excel"``, ``"html"``.

    read_kwargs : dict or None
        Extra keyword arguments passed to the pandas reader function.

    html_table_index : int
        Which HTML table to use when ``read_format="html"`` (default 0).

    transform_fn : callable or None
        Optional ``transform_fn(df) -> df`` for custom post-processing
        (e.g. rename French columns, convert dates, strip whitespace).

    column_names : dict or None
        Rename columns: ``{"OldName": "NewName"}``.

    column_order : list or None
        Reorder columns: ``["Col1", "Col2", ...]``.
        Columns not in the list are dropped.

    upload_to_sharepoint : bool
        True = upload to SharePoint via ``save_excel_to_sharepoint()``.

    template_path : str
        Full path to the Excel template (e.g. ``TEMPLATE_DO_NOT_DELETE.xlsx``).

    sharepoint_folder : str
        SharePoint destination folder (mapped drive path).

    output_filename_prefix : str
        Base filename for the SharePoint file (without .xlsx).

    output_use_date : bool
        Append today's date to the output filename (default True).

    output_use_time : bool
        Append current time to the output filename (default False).

    column_types : dict or None
        Column formatting for the SharePoint Excel file.
        E.g. ``{"Amount": "currency", "Date": "date"}``.
        Types: number, currency, date, time, percentage, fraction, text, general.

    delete_after_upload : bool
        True = delete the local downloaded file after upload.

    Returns
    -------
    pd.DataFrame
        The loaded (and optionally transformed) data.
    """
    if download_dir is None:
        download_dir = os.path.join(os.environ["USERPROFILE"], "Downloads")

    file_path = None
    df = None

    # --- Step 1: Launch browser -----------------------------------------
    print("Step 1: Launching browser...")
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=headless)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # --- Step 2: Navigate to URL ------------------------------------
        print(f"Step 2: Navigating to {url}")
        page.goto(url, wait_until=wait_until)

        # --- Step 3: Run user's scraping function -----------------------
        print("Step 3: Running scrape function...")
        file_path = scrape_fn(page, download_dir)
        print(f"  Downloaded: {file_path}")

        # --- Step 4: Close browser --------------------------------------
        browser.close()
        print("Step 4: Browser closed.")

    # --- Step 5: Read into DataFrame ------------------------------------
    print("Step 5: Reading downloaded file...")
    df = _read_file(file_path, read_format=read_format,
                    read_kwargs=read_kwargs, html_table_index=html_table_index)
    print(f"  Data loaded: {df.shape[0]} rows x {df.shape[1]} columns")

    # --- Step 6: Post-processing ----------------------------------------
    if transform_fn is not None:
        print("Step 6: Applying transform...")
        df = transform_fn(df)
        print(f"  After transform: {df.shape[0]} rows x {df.shape[1]} columns")

    if column_names:
        df.rename(columns=column_names, inplace=True)
        print(f"  Renamed columns: {list(column_names.values())}")

    if column_order:
        cols_present = [c for c in column_order if c in df.columns]
        df = df[cols_present]
        print(f"  Reordered to {len(cols_present)} columns")

    # --- Step 7: SharePoint upload --------------------------------------
    if upload_to_sharepoint:
        prefix = output_filename_prefix or "WebScrape"
        sp_prefix = build_filename(prefix, output_use_date, output_use_time)

        print(f"Step 7: Uploading to SharePoint as {sp_prefix}.xlsx ...")
        result = save_excel_to_sharepoint(
            df,
            template_path=template_path,
            sharepoint_folder=sharepoint_folder,
            output_filename_prefix=sp_prefix,
            column_types=column_types,
        )
        print(f"  Uploaded to: {result['sharepoint_path']}")
    else:
        print("Step 7: SharePoint upload skipped (upload_to_sharepoint=False).")

    # --- Step 8: Cleanup ------------------------------------------------
    if delete_after_upload and file_path and os.path.exists(file_path):
        try:
            os.remove(file_path)
            print(f"Step 8: Cleaned up download: {file_path}")
        except Exception as e:
            print(f"Step 8: [Warning] Could not delete download: {e}")

    print("\nDone!")
    return df
