# SAP BW MDX Export
# =================
# Query SAP BW (BEx queries / InfoProviders) via MDX over the OLAP BAPIs and
# export the result to CSV or XLSX. Mirrors the surface of sap_rfc_export.py.
#
# High-level entry point:
#   - run_bw_mdx_extract(config)  -> one-call BW export driven by a config dict
#
# Public functions:
#   - bw_query()        -> returns a DataFrame (own session)
#   - bw_query_conn()   -> returns a DataFrame using an open connection
#   - bw_export()       -> save DataFrame to CSV/XLSX, returns file path
#
# Notes:
#   - The target BEx query must be released for OLE DB for OLAP (RSRT).
#   - Variables are passed via the standard "SAP VARIABLES" MDX clause.

import pandas as pd
from contextlib import contextmanager

from .sap_rfc_connection import RFCConnection, WinCredentialStore
from .sap_rfc_export import save_df, build_output_name


# -- Connection helper ---------------------------------------------------------

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


# -- MDX assembly --------------------------------------------------------------

def _build_mdx(query: str,
               rows: list[str],
               columns: list[str],
               where: str = None,
               variables: dict = None) -> str:
    """
    Assemble a simple MDX statement for a BEx query.

    rows / columns are sets of MDX member expressions, e.g.
        ["[0MATERIAL].[LEVEL01].MEMBERS"]
        ["[Measures].[ZSALES]", "[Measures].[ZQTY]"]
    """
    cols_set = "{" + ", ".join(columns) + "}"
    rows_set = "{" + ", ".join(rows) + "}"
    cube = f"[QUERYCUBE/{query}]"

    mdx = f"SELECT {cols_set} ON COLUMNS, {rows_set} ON ROWS FROM {cube}"
    if where:
        mdx += f" WHERE {where}"

    if variables:
        parts = []
        for var, val in variables.items():
            parts.append(f"[{var}] INCLUDING '{val}'")
        mdx += " SAP VARIABLES " + " ".join(parts)

    return mdx


# -- Result flattening ---------------------------------------------------------

def _tuple_caption(members: list[dict]) -> str:
    """Join MEMBER_CAPTIONs of a tuple with ' | '."""
    return " | ".join(m.get("member_caption", "") for m in members)


def _mdx_result_to_df(raw: dict) -> pd.DataFrame:
    """
    Flatten the dict returned by RFCConnection.bapi_mdx_query into a DataFrame.

    Layout assumption: axis 0 = COLUMNS, axis 1 = ROWS (standard MDX).
    Columns of the resulting DataFrame:
        - one leading column per dimension on the row axis
          (named after that tuple's first member's dimension)
        - one column per tuple on the column axis (caption-joined)
    """
    axes = raw.get("axes", [])
    cells = raw.get("cells", [])

    if len(axes) == 0:
        return pd.DataFrame()

    col_axis = axes[0]
    row_axis = axes[1] if len(axes) > 1 else {"tuples": [[]]}

    col_tuples = col_axis["tuples"]
    row_tuples = row_axis["tuples"]

    n_cols = max(len(col_tuples), 1)
    n_rows = max(len(row_tuples), 1)

    # Row dimension column names — derived from the first row tuple
    row_dim_names = []
    if row_tuples and row_tuples[0]:
        for m in row_tuples[0]:
            name = m.get("dimension_unique_name") or m.get("hierarchy_unique_name") or "DIM"
            row_dim_names.append(name)

    # Column headers
    col_headers = [_tuple_caption(t) or f"COL_{i}" for i, t in enumerate(col_tuples)]

    # Build empty grid
    data = {h: [None] * n_rows for h in col_headers}

    # Place cell values
    for c in cells:
        ord_ = c["cell_ordinal"]
        col_idx = ord_ % n_cols
        row_idx = ord_ // n_cols
        if row_idx >= n_rows or col_idx >= n_cols:
            continue
        val = c.get("value")
        if val == "" or val is None:
            val = c.get("formatted_value")
        try:
            val = float(val)
        except (TypeError, ValueError):
            pass
        data[col_headers[col_idx]][row_idx] = val

    df = pd.DataFrame(data)

    # Prepend row dimension columns
    for dim_idx, dim_name in enumerate(row_dim_names):
        df.insert(
            dim_idx,
            dim_name,
            [(_member_caption(row_tuples[r], dim_idx) if r < len(row_tuples) else "")
             for r in range(n_rows)],
        )

    return df


def _member_caption(tuple_members: list[dict], dim_idx: int) -> str:
    if dim_idx < len(tuple_members):
        return tuple_members[dim_idx].get("member_caption", "")
    return ""


# -- Public query functions ----------------------------------------------------

def bw_query_conn(conn, mdx: str) -> pd.DataFrame:
    """Run an MDX query using an already-open connection. Returns a DataFrame."""
    print(f"\nMDX:\n{mdx}\n")
    raw = conn.bapi_mdx_query(mdx)
    df = _mdx_result_to_df(raw)
    print(f"Query complete: {df.shape[0]:,} rows x {df.shape[1]} cols")
    return df


def bw_query(mdx: str) -> pd.DataFrame:
    """Run an MDX query in its own SAP session. Returns a DataFrame."""
    with _sap_connection() as conn:
        return bw_query_conn(conn, mdx)


def bw_export(mdx: str,
              output_dir: str = ".",
              output_name: str = "bw_export",
              extension: str = "csv") -> str:
    """Run an MDX query and save the result to CSV or XLSX. Returns the file path."""
    df = bw_query(mdx)
    return save_df(df, output_dir, output_name, extension)


# -- High-level entry point ----------------------------------------------------

def run_bw_mdx_extract(config: dict) -> str:
    """
    One-call BW export driven by a config dict.

    Required keys:
        output_dir
        AND either:
            mdx
        OR:
            query, rows, columns

    Optional keys:
        where, variables, output_name, extension, add_date, add_time
    """
    output_dir = config["output_dir"]

    extension = config.get("extension", "csv")
    add_date = config.get("add_date", False)
    add_time = config.get("add_time", False)
    output_name = config.get("output_name")

    mdx = config.get("mdx")
    if not mdx:
        query = config.get("query")
        rows = config.get("rows")
        columns = config.get("columns")
        if not (query and rows and columns):
            raise ValueError(
                "run_bw_mdx_extract: provide either 'mdx' or all of 'query', 'rows', 'columns'."
            )
        mdx = _build_mdx(
            query=query,
            rows=rows,
            columns=columns,
            where=config.get("where"),
            variables=config.get("variables"),
        )

    base = output_name or (config.get("query") or "bw_export")
    name = build_output_name(base, add_date, add_time)
    print(f"Output name: {name}.{extension}")

    df = bw_query(mdx)
    return save_df(df, output_dir, name, extension)
