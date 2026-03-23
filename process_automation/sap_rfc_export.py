# SAP RFC Export
# ==============
# Paginated SAP table query and export via RFC.
#
# High-level entry point:
#   - run_rfc_extract(config)  -> one-call export driven by a config dict
#
# Public query functions:
#   - sap_query()              -> returns a DataFrame (one SAP session, no file saved)
#   - sap_export()             -> saves result to CSV or XLSX, returns file path
#   - sap_chained_export()     -> two-step query: lookup table -> chunked main export
#   - sap_query_conn()         -> low-level: fetch using an existing open connection
#
# Schema functions:
#   - get_table_schema()       -> return schema DataFrame for a SAP table
#   - save_table_schema()      -> get_table_schema() + save to file
#
# Utility helpers:
#   - save_df()                -> save a DataFrame to CSV or XLSX
#   - build_output_name()      -> append date/time suffix to a filename
#   - build_filter_from_file() -> build OR filter from a CSV/XLSX column
#   - chunk_values()           -> split a list into N-sized chunks
#   - build_filter_chunk()     -> build OR filter from a list of values
#
# All query functions paginate automatically via rowcount/rowskips.
#
# Usage:
#   from process_automation import run_rfc_extract
#   from process_automation import sap_query, sap_export

import os
import pandas as pd
from pathlib import Path
from datetime import datetime
from contextlib import contextmanager
from .sap_rfc_connection import RFCConnection, WinCredentialStore

BATCH_SIZE = 10_000
COL_BATCH_SIZE = 30


# -- Internal helpers ----------------------------------------------------------

def _fetch_column_batches(conn, table: str, cols: str, filters: list[str], batch_size: int, max_pages: int = 0, col_batch_size: int = COL_BATCH_SIZE) -> pd.DataFrame:
    """
    Fetch a table in column batches to avoid the 512-char row width limit
    of BBP_RFC_READ_TABLE. Each batch fetches a subset of columns with full
    row pagination, then all batches are merged horizontally by row index.
    """
    col_list = [c.strip() for c in cols.split(",")]

    if col_batch_size <= 0 or len(col_list) <= col_batch_size:
        return _fetch_all_pages(conn, table, cols, filters, batch_size, max_pages)

    col_batches = [col_list[i:i + col_batch_size] for i in range(0, len(col_list), col_batch_size)]
    print(f"Splitting {len(col_list)} columns into {len(col_batches)} batch(es) of up to {col_batch_size}")

    dfs = []
    for i, batch in enumerate(col_batches, 1):
        batch_cols = ",".join(batch)
        print(f"\n   Column batch {i}/{len(col_batches)}  ({len(batch)} cols: {batch_cols[:60]}{'...' if len(batch_cols) > 60 else ''})")
        df_batch = _fetch_all_pages(conn, table, batch_cols, filters, batch_size, max_pages)
        if df_batch.empty and i == 1:
            return pd.DataFrame()
        dfs.append(df_batch)

    df = pd.concat(dfs, axis=1)
    # Remove duplicate columns (if any overlap)
    df = df.loc[:, ~df.columns.duplicated()]
    print(f"\nMerged {len(col_batches)} column batches: {df.shape[0]:,} rows x {df.shape[1]} cols")
    return df


def _fetch_all_pages(conn, table: str, cols: str, filters: list[str], batch_size: int, max_pages: int = 0) -> pd.DataFrame:
    """
    Call bbp_rfc_read_table repeatedly until all rows are retrieved.
    Returns a single concatenated DataFrame.
    None from the RFC = valid empty result (zero rows), not an error.
    max_pages: if > 0, stop after that many pages (useful for debug).
    """
    pages = []
    offset = 0
    page_num = 0

    while True:
        page_num += 1
        print(f"Fetching rows {offset + 1} – {offset + batch_size}  (page {page_num})...")

        data = conn.bbp_rfc_read_table(
            table,
            cols,
            filters,
            rowcount=batch_size,
            rowskips=offset,
        )

        if data is None:
            print("   -> No rows returned — done.")
            break

        df_page = pd.DataFrame(data)

        if df_page.empty:
            print("   -> Empty page — done.")
            break

        pages.append(df_page)
        fetched = len(df_page)
        print(f"   -> {fetched:,} rows received.")

        if fetched < batch_size:
            break

        if max_pages > 0 and page_num >= max_pages:
            print(f"   -> max_pages={max_pages} reached — stopping early.")
            break

        offset += batch_size

    df = pd.concat(pages, ignore_index=True) if pages else pd.DataFrame()
    total = len(df)
    print(f"Query complete: {total:,} rows total across {page_num} page(s)")
    return df


@contextmanager
def _sap_connection():
    """Open a SAP connection and ensure it is closed on exit."""
    WinCredentialStore.ensure_credentials()
    conn = RFCConnection()
    print("\nLogging in to SAP...")
    if conn.login() is None:
        raise RuntimeError("SAP login failed.")
    try:
        yield conn
    finally:
        try:
            conn.log_off()
            print("Logged off SAP")
        except Exception:
            pass


# -- Utility helpers -----------------------------------------------------------

def save_df(df: pd.DataFrame, output_dir: str, output_name: str, extension: str = "csv") -> str:
    """Save a DataFrame to CSV or XLSX. Returns the file path."""
    extension = extension.lower().lstrip(".")
    if extension not in ("csv", "xlsx"):
        raise ValueError(f"Unsupported extension '{extension}'. Use 'csv' or 'xlsx'.")

    final_path = os.path.join(output_dir, f"{output_name}.{extension}")

    if os.path.exists(final_path):
        os.remove(final_path)

    if extension == "csv":
        df.to_csv(final_path, index=False, encoding="utf-8-sig")
    else:
        df.to_excel(final_path, index=False, engine="openpyxl")

    print(f"Saved: {final_path}")
    return final_path


def build_output_name(output_name: str, add_date: bool = False, add_time: bool = False) -> str:
    """Append _YYYYMMDD and/or _HHMMSS to output_name."""
    now = datetime.now()
    suffix = ""
    if add_date:
        suffix += f"_{now.strftime('%Y%m%d')}"
    if add_time:
        suffix += f"_{now.strftime('%H%M%S')}"
    return f"{output_name}{suffix}"


def build_filter_from_file(file_path: str, column: str, sap_field: str) -> list[str]:
    """Read a CSV/XLSX file and return an OR filter from a column's unique values."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {file_path}")

    print(f"Reading filter file: {file_path}")
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=str)
    elif path.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(path, dtype=str)
    else:
        raise ValueError(f"Unsupported file type: {path.suffix}. Use .csv or .xlsx.")

    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found. Available: {list(df.columns)}")

    values = df[column].dropna().unique().tolist()
    if not values:
        raise ValueError(f"Column '{column}' is empty — no values to filter on.")

    print(f"   -> {len(values):,} unique value(s) found in '{column}'")
    conditions = " OR ".join(f"{sap_field} = '{v.strip()}'" for v in values)
    return [f"({conditions})"]


def chunk_values(values: list, chunk_size: int) -> list[list]:
    """Split a list into chunks of at most chunk_size items."""
    return [values[i:i + chunk_size] for i in range(0, len(values), chunk_size)]


def build_filter_chunk(values: list, sap_field: str) -> list[str]:
    """Build a single OR filter string from a chunk of values."""
    conditions = " OR ".join(f"{sap_field} = '{str(v).strip()}'" for v in values)
    return [f"({conditions})"]


# -- Public query functions ----------------------------------------------------

def sap_query_conn(
    conn,
    table: str,
    cols: str | list[str],
    filters: list[str],
    batch_size: int = BATCH_SIZE,
) -> pd.DataFrame:
    """
    Fetch a SAP table using an already-open connection. Returns a DataFrame.
    Use this inside chunk loops to avoid reconnecting on every chunk.
    """
    if isinstance(cols, list):
        cols = ",".join(cols)
    return _fetch_column_batches(conn, table, cols, filters, batch_size)


def sap_query(
    table: str,
    cols: str | list[str],
    filters: list[str],
    batch_size: int = BATCH_SIZE,
    max_pages: int = 0,
) -> pd.DataFrame:
    """
    Fetch a SAP table and return a DataFrame. Opens and closes its own session.
    No file is written. Useful for chaining results into a second call.
    max_pages: if > 0, stop after that many pages (e.g. max_pages=1 for a quick sample).
    """
    if isinstance(cols, list):
        cols = ",".join(cols)

    with _sap_connection() as conn:
        print(f"   [{table}]")
        return _fetch_column_batches(conn, table, cols, filters, batch_size, max_pages=max_pages)


def sap_export(
    table: str,
    cols: str | list[str],
    filters: list[str],
    output_dir: str = ".",
    output_name: str = None,
    extension: str = "csv",
    batch_size: int = BATCH_SIZE,
) -> str:
    """
    Fetch a SAP table and save to CSV or XLSX. Returns the file path.
    Opens and closes its own SAP session.
    """
    if isinstance(cols, list):
        cols = ",".join(cols)

    with _sap_connection() as conn:
        print(f"   [{table}]")
        df = _fetch_column_batches(conn, table, cols, filters, batch_size)

    return save_df(df, output_dir, output_name or table, extension)


def sap_chained_export(
    # Step 1 — lookup query
    chain_table: str,
    chain_cols: str | list[str],
    chain_filters: list[str],
    chain_source_column: str,
    # Step 2 — main export
    table: str,
    cols: str | list[str],
    target_field: str,
    # Output
    output_dir: str,
    output_name: str,
    extension: str = "csv",
    # Tuning
    batch_size: int = BATCH_SIZE,
    chain_batch_size: int = BATCH_SIZE,
    chunk_size: int = 2,
    # Optional: save Step 1 result
    save_chain_file: bool = False,
    chain_output_dir: str = None,
    chain_output_name: str = None,
    chain_extension: str = "csv",
) -> str:
    """
    Two-step chained export: query one SAP table, use a column from the
    result as a chunked filter for a second table. Returns the final file path.

    A single SAP connection is reused across all Step 2 chunks.
    """
    # Step 1 — fetch the lookup table
    print(f"\nStep 1 — querying {chain_table} to build filter...")
    df_chain = sap_query(
        table=chain_table,
        cols=chain_cols,
        filters=chain_filters,
        batch_size=chain_batch_size,
    )

    if save_chain_file and chain_output_dir and chain_output_name:
        save_df(df_chain, chain_output_dir, chain_output_name, chain_extension)

    if chain_source_column not in df_chain.columns:
        raise ValueError(
            f"Column '{chain_source_column}' not found in {chain_table} result. "
            f"Available: {list(df_chain.columns)}"
        )

    values = df_chain[chain_source_column].dropna().unique().tolist()
    if not values:
        raise ValueError(f"No values found in '{chain_source_column}' — cannot build filter.")

    chunks = chunk_values(values, chunk_size)
    print(f"\nStep 2 — exporting {table} in {len(chunks)} chunk(s) "
          f"({len(values):,} unique {target_field} values, {chunk_size} per chunk)...")

    # Step 2 — one connection reused across all chunks
    if isinstance(cols, list):
        cols = ",".join(cols)

    dfs = []
    with _sap_connection() as conn:
        print(f"   [{table}]")

        for i, chunk in enumerate(chunks, 1):
            chunk_filter = build_filter_chunk(chunk, target_field)
            print(f"\n   Chunk {i}/{len(chunks)}  ({len(chunk)} values)  "
                  f"filter: {chunk_filter[0][:80]}{'...' if len(chunk_filter[0]) > 80 else ''}")

            df_chunk = _fetch_column_batches(conn, table, cols, chunk_filter, batch_size)
            if not df_chunk.empty:
                dfs.append(df_chunk)

    df_final = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    print(f"\nAll chunks done: {len(df_final):,} total rows")

    return save_df(df_final, output_dir, output_name, extension)


# -- Schema functions ----------------------------------------------------------

def get_table_schema(table: str, field_names: str = "both", show_example_row: bool = True) -> pd.DataFrame:
    """
    Fetch field metadata for a SAP table using DD03L + DD04T via bbp_rfc_read_table.

    DD03L : field definitions (name, datatype, length, key flag, rollname)
    DD04T : data element texts (short/medium/long labels) — joined on ROLLNAME

    Parameters
    ----------
    table           : SAP table name
    field_names     : "technical", "label", or "both"
    show_example_row: if True, fetch one real row and add EXAMPLE column

    Returns
    -------
    pd.DataFrame
    """
    field_names = field_names.lower()
    if field_names not in ("technical", "label", "both"):
        raise ValueError(f"Invalid field_names '{field_names}'. Use 'technical', 'label', or 'both'.")

    # Step 1: field definitions from DD03L
    print(f"\nReading field definitions from DD03L for table: {table}...")
    df_dd03l = sap_query(
        table="DD03L",
        cols="TABNAME,FIELDNAME,POSITION,KEYFLAG,ROLLNAME,DATATYPE,LENG",
        filters=[f"TABNAME = '{table}'"],
    )

    if df_dd03l.empty:
        raise RuntimeError(f"No fields found for table '{table}'. Check the table name.")

    df_dd03l.columns = ["TABNAME", "FIELDNAME", "POSITION", "KEYFLAG", "ROLLNAME", "DATATYPE", "LENG"]

    # Keep only real fields (exclude .INCLUDE and empty fieldnames)
    df_dd03l = df_dd03l[
        df_dd03l["FIELDNAME"].str.strip().ne("") &
        ~df_dd03l["FIELDNAME"].str.startswith(".")
    ].copy()

    print(f"   -> {len(df_dd03l)} fields found.")

    # Step 2: labels from DD04T (joined on ROLLNAME)
    label_map = {}
    if field_names in ("label", "both"):
        rollnames = df_dd03l["ROLLNAME"].dropna().unique().tolist()
        rollnames = [r for r in rollnames if r.strip()]

        if rollnames:
            print(f"   -> Fetching labels from DD04T for {len(rollnames)} data elements...")
            conditions = " OR ".join(f"ROLLNAME = '{r}'" for r in rollnames)
            df_dd04t = sap_query(
                table="DD04T",
                cols="ROLLNAME,DDLANGUAGE,DDTEXT,REPTEXT,SCRTEXT_S,SCRTEXT_M,SCRTEXT_L",
                filters=[
                    f"({conditions})",
                    "DDLANGUAGE = 'E'",
                ],
            )
            if not df_dd04t.empty:
                df_dd04t.columns = ["ROLLNAME", "DDLANGUAGE", "DDTEXT", "REPTEXT", "SCRTEXT_S", "SCRTEXT_M", "SCRTEXT_L"]
                df_dd04t["LABEL"] = (
                    df_dd04t["DDTEXT"].str.strip()
                    .where(df_dd04t["DDTEXT"].str.strip().ne(""), other=None)
                    .fillna(df_dd04t["SCRTEXT_L"].str.strip())
                    .where(df_dd04t["SCRTEXT_L"].str.strip().ne(""), other=None)
                    .fillna(df_dd04t["SCRTEXT_M"].str.strip())
                    .fillna(df_dd04t["SCRTEXT_S"].str.strip())
                    .fillna(df_dd04t["REPTEXT"].str.strip())
                    .fillna("")
                )
                label_map = df_dd04t.set_index("ROLLNAME")["LABEL"].to_dict()

        df_dd03l["LABEL"] = df_dd03l["ROLLNAME"].map(label_map).fillna("")

    # Step 3: example row from the actual table (no filter)
    example_map = {}
    if show_example_row:
        field_list = ",".join(df_dd03l["FIELDNAME"].str.strip().tolist())
        print(f"   -> Fetching one example row from {table}...")
        df_example = sap_query(
            table=table,
            cols=field_list,
            filters=["MANDT >= '000'"],
            batch_size=BATCH_SIZE,
            max_pages=1,
        )
        df_example = df_example.head(1)
        if not df_example.empty:
            fieldnames = df_dd03l["FIELDNAME"].str.strip().tolist()
            row = df_example.iloc[0]
            for i, fname in enumerate(fieldnames):
                example_map[fname] = row.iloc[i] if i < len(row) else ""
            print(f"   -> Example row retrieved.")
        else:
            print(f"   -> No example row found (table may be empty or filter too strict).")

        df_dd03l["EXAMPLE"] = df_dd03l["FIELDNAME"].str.strip().map(example_map).fillna("")

    # Build final DataFrame
    rename_map = {"LENG": "LENGTH", "KEYFLAG": "KEY"}

    if field_names == "technical":
        cols = ["POSITION", "FIELDNAME", "DATATYPE", "LENG", "KEYFLAG"]
    elif field_names == "label":
        cols = ["POSITION", "LABEL", "DATATYPE", "LENG", "KEYFLAG"]
    else:  # both
        cols = ["POSITION", "FIELDNAME", "LABEL", "DATATYPE", "LENG", "KEYFLAG"]

    if show_example_row:
        cols.append("EXAMPLE")

    df = df_dd03l[cols].copy()
    df = df.rename(columns=rename_map)
    df = df.sort_values("POSITION").reset_index(drop=True)

    return df


def save_table_schema(
    table: str,
    output_dir: str = ".",
    extension: str = "csv",
    field_names: str = "both",
    show_example_row: bool = True,
    add_date: bool = False,
    add_time: bool = False,
    output_name: str = None,
) -> str:
    """Fetch table schema and save to file. Returns the file path."""
    name = build_output_name(output_name or f"{table}_schema", add_date, add_time)
    print(f"Output name: {name}.{extension}")

    df = get_table_schema(table, field_names, show_example_row)
    print("\n" + df.to_string(index=False))

    return save_df(df, output_dir, name, extension)


# -- High-level entry point ----------------------------------------------------

def run_rfc_extract(config: dict) -> str:
    """
    One-call SAP export driven by a config dict.

    Required keys:
        table, cols, output_dir

    Optional keys:
        filters, extension, add_date, add_time, output_name, batch_size,
        file_filter, chain

    Returns the output file path.
    """
    # Required keys
    table = config["table"]
    cols = config["cols"]
    output_dir = config["output_dir"]

    # Optional keys
    filters = config.get("filters", [])
    extension = config.get("extension", "csv")
    add_date = config.get("add_date", False)
    add_time = config.get("add_time", False)
    output_name = config.get("output_name")
    batch_size = config.get("batch_size", BATCH_SIZE)
    file_filter = config.get("file_filter")
    chain = config.get("chain")

    # Validate
    if file_filter and chain:
        raise ValueError("Cannot use both 'file_filter' and 'chain' at the same time. Pick one.")

    # Build output name
    name = build_output_name(output_name or table, add_date, add_time)
    print(f"Output name: {name}.{extension}")

    # Mode: chained two-step export
    if chain:
        return sap_chained_export(
            chain_table=chain["table"],
            chain_cols=chain["cols"],
            chain_filters=chain["filters"],
            chain_source_column=chain["source_column"],
            table=table,
            cols=cols,
            target_field=chain["target_field"],
            output_dir=output_dir,
            output_name=name,
            extension=extension,
            batch_size=batch_size,
            chain_batch_size=chain.get("batch_size", BATCH_SIZE),
            chunk_size=chain.get("chunk_size", 2),
            save_chain_file=chain.get("save_file", False),
            chain_output_dir=chain.get("save_dir"),
            chain_output_name=chain.get("save_name"),
            chain_extension=chain.get("save_extension", "csv"),
        )

    # Mode: file-based filter
    if file_filter:
        filters = build_filter_from_file(
            file_path=file_filter["path"],
            column=file_filter["column"],
            sap_field=file_filter["field"],
        )

    # Mode: plain filters (default)
    return sap_export(
        table=table,
        cols=cols,
        filters=filters,
        output_dir=output_dir,
        output_name=name,
        extension=extension,
        batch_size=batch_size,
    )
