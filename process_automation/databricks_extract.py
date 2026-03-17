# Databricks Extract Engine
# =========================
# Pulls data from a Databricks table (or custom query), then saves locally
# and/or uploads to SharePoint.
#
# Usage:
#   from process_automation import run_databricks_extract
#
#   df = run_databricks_extract(
#       table_path="catalog.schema.table_name",
#       select_columns=["col1", "col2"],
#       filters=["status = 'ACTIVE'"],
#       limit=10000,
#       upload_to_sharepoint=True,
#       template_path=r"Z:\path\to\TEMPLATE_DO_NOT_DELETE.xlsx",
#       sharepoint_folder=r"Z:\path\to\folder",
#   )

import os
import re
from typing import Dict, Optional, Sequence, Union

import pandas as pd

from .sap_extract import build_filename
from .sharepoint_upload import save_excel_to_sharepoint

_IDENTIFIER_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_BOOL_TRUE = {"1", "true", "t", "yes", "y"}
_BOOL_FALSE = {"0", "false", "f", "no", "n"}


def _load_env_file(env_path: Optional[str]) -> None:
    if not env_path or not os.path.isfile(env_path):
        return

    with open(env_path, "r", encoding="utf-8") as handle:
        for raw_line in handle:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue

            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip("'").strip('"')
            if key and key not in os.environ:
                os.environ[key] = value


def _coerce_bool(value, default: Optional[bool] = None) -> bool:
    if value is None:
        if default is None:
            raise ValueError("Boolean setting cannot be None.")
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized == "default":
            if default is None:
                raise ValueError("Boolean setting cannot be 'Default' without a fallback value.")
            return default
        if normalized in _BOOL_TRUE:
            return True
        if normalized in _BOOL_FALSE:
            return False
    raise ValueError(f"Invalid boolean setting: {value}")


def _quote_identifier(identifier: str) -> str:
    token = identifier.strip()
    if not token:
        raise ValueError("Identifier cannot be empty.")

    if token.startswith("`") and token.endswith("`"):
        inner = token[1:-1].strip()
        if not inner:
            raise ValueError("Identifier cannot be empty inside backticks.")
        return token

    if not _IDENTIFIER_RE.fullmatch(token):
        raise ValueError(
            f"Unsafe identifier '{identifier}'. Use letters, numbers, and underscores, "
            "or pass a custom query."
        )
    return f"`{token}`"


def _quote_table_path(table_path: str) -> str:
    parts = [part.strip() for part in table_path.split(".") if part.strip()]
    if not parts:
        raise ValueError("table_path is required when query is not provided.")
    return ".".join(_quote_identifier(part) for part in parts)


def _format_select_column(column_expr: str) -> str:
    expr = column_expr.strip()
    if not expr:
        raise ValueError("select_columns contains an empty value.")
    if expr == "*":
        return expr

    parts = [part.strip() for part in expr.split(".")]
    if all(_IDENTIFIER_RE.fullmatch(part) for part in parts):
        return ".".join(_quote_identifier(part) for part in parts)
    return expr


def _build_query(
    table_path: str,
    query: Optional[str] = None,
    select_columns: Optional[Sequence[str]] = None,
    filters: Optional[Sequence[str]] = None,
    order_by: Optional[Union[str, Sequence[str]]] = None,
    limit: Optional[int] = None,
) -> str:
    if query:
        return query.strip()

    select_sql = "*"
    if select_columns:
        select_sql = ", ".join(_format_select_column(col) for col in select_columns)

    sql = [f"SELECT {select_sql}", f"FROM {_quote_table_path(table_path)}"]

    valid_filters = [clause.strip() for clause in (filters or []) if clause and clause.strip()]
    if valid_filters:
        sql.append("WHERE " + " AND ".join(valid_filters))

    if isinstance(order_by, str):
        order_clause = order_by.strip()
        if order_clause:
            sql.append(f"ORDER BY {order_clause}")
    elif order_by:
        clauses = [clause.strip() for clause in order_by if clause and clause.strip()]
        if clauses:
            sql.append("ORDER BY " + ", ".join(clauses))

    if limit is not None:
        if limit <= 0:
            raise ValueError("limit must be greater than 0.")
        sql.append(f"LIMIT {int(limit)}")

    return "\n".join(sql)


def _save_local_dataframe(
    df: pd.DataFrame,
    output_dir: str,
    output_filename_prefix: str,
    output_format: str,
) -> str:
    fmt = output_format.strip().lower()
    if fmt == "xlsx":
        output_path = os.path.join(output_dir, f"{output_filename_prefix}.xlsx")
        df.to_excel(output_path, index=False)
    elif fmt == "csv":
        output_path = os.path.join(output_dir, f"{output_filename_prefix}.csv")
        df.to_csv(output_path, index=False)
    elif fmt == "parquet":
        output_path = os.path.join(output_dir, f"{output_filename_prefix}.parquet")
        df.to_parquet(output_path, index=False)
    else:
        raise ValueError("output_format must be one of: 'xlsx', 'csv', 'parquet'")
    return output_path


def _resolve_databricks_config(env_path: Optional[str]) -> Dict[str, str]:
    default_env = os.path.join(os.path.dirname(__file__), ".env")
    _load_env_file(env_path or default_env)

    host = (
        os.getenv("DATABRICKS_SERVER_HOSTNAME")
        or os.getenv("DATABRICKS_HOST")
        or ""
    ).strip()
    token = (os.getenv("DATABRICKS_TOKEN") or "").strip()
    http_path = (os.getenv("DATABRICKS_HTTP_PATH") or "").strip()

    if host.startswith("https://"):
        host = host[len("https://"):]
    if host.startswith("http://"):
        host = host[len("http://"):]
    host = host.rstrip("/")

    if not http_path:
        cluster_id = (os.getenv("DATABRICKS_CLUSTER_ID") or "").strip()
        org_id = (os.getenv("DATABRICKS_ORG_ID") or os.getenv("DATABRICKS_WORKSPACE_ID") or "").strip()
        if cluster_id and org_id:
            http_path = f"/sql/protocolv1/o/{org_id}/{cluster_id}"

    if not host:
        raise ValueError("Missing Databricks host. Set DATABRICKS_HOST or DATABRICKS_SERVER_HOSTNAME.")
    if not token:
        raise ValueError("Missing Databricks token. Set DATABRICKS_TOKEN.")
    if not http_path:
        raise ValueError(
            "Missing Databricks HTTP path. Set DATABRICKS_HTTP_PATH, or set both "
            "DATABRICKS_ORG_ID and DATABRICKS_CLUSTER_ID."
        )

    return {
        "server_hostname": host,
        "access_token": token,
        "http_path": http_path,
    }


def run_databricks_extract(
    table_path: str = "",
    query: Optional[str] = None,
    env_path: Optional[str] = None,
    select_columns: Optional[Sequence[str]] = None,
    filters: Optional[Sequence[str]] = None,
    order_by: Optional[Union[str, Sequence[str]]] = None,
    limit: Optional[int] = None,
    save_local: bool = True,
    output_format: str = "xlsx",
    download_dir: Optional[str] = None,
    download_filename: str = "",
    download_use_date: bool = True,
    download_use_time: bool = True,
    upload_to_sharepoint: bool = False,
    template_path: str = "",
    sharepoint_folder: str = "",
    sharepoint_filename: str = "Default",
    sharepoint_use_date: Union[str, bool] = "Default",
    sharepoint_use_time: Union[str, bool] = "Default",
    keep_desktop_copy: bool = False,
    excel_visible: bool = False,
    column_types: Optional[Dict[str, str]] = None,
) -> pd.DataFrame:
    """Extract data from Databricks and save locally and/or upload to SharePoint."""

    if not query and not table_path:
        raise ValueError("Provide either table_path or query.")
    if not save_local and not upload_to_sharepoint:
        raise ValueError("At least one output must be enabled: save_local or upload_to_sharepoint.")
    if upload_to_sharepoint and (not template_path or not sharepoint_folder):
        raise ValueError("template_path and sharepoint_folder are required when upload_to_sharepoint=True.")

    print("=" * 60)
    print("DATABRICKS EXTRACT - STARTED")
    print("=" * 60)

    config = _resolve_databricks_config(env_path)
    sql_text = _build_query(
        table_path=table_path,
        query=query,
        select_columns=select_columns,
        filters=filters,
        order_by=order_by,
        limit=limit,
    )
    print("[OK] Query prepared")

    try:
        from databricks import sql as databricks_sql
    except ImportError as exc:
        raise ImportError(
            "Missing dependency 'databricks-sql-connector'. Install it with:\n"
            "pip install databricks-sql-connector"
        ) from exc

    print("[Run] Executing query on Databricks...")
    with databricks_sql.connect(
        server_hostname=config["server_hostname"],
        http_path=config["http_path"],
        access_token=config["access_token"],
    ) as connection:
        with connection.cursor() as cursor:
            cursor.execute(sql_text)
            rows = cursor.fetchall()
            columns = [desc[0] for desc in cursor.description]
            df_result = pd.DataFrame(rows, columns=columns)

    print(f"[OK] Retrieved {len(df_result)} rows x {len(df_result.columns)} cols")

    default_base = download_filename or (table_path.split(".")[-1] if table_path else "databricks_extract")
    local_prefix = build_filename(default_base, download_use_date, download_use_time)

    local_output_path = None
    if save_local:
        if download_dir is None:
            download_dir = os.path.join(os.environ["USERPROFILE"], "Downloads")
        os.makedirs(download_dir, exist_ok=True)
        local_output_path = _save_local_dataframe(df_result, download_dir, local_prefix, output_format)
        print(f"[OK] Local file saved: {local_output_path}")

    sharepoint_output_path = None
    if upload_to_sharepoint:
        sp_base = default_base if str(sharepoint_filename).strip().lower() == "default" else str(sharepoint_filename).strip()
        sp_use_date = _coerce_bool(sharepoint_use_date, default=download_use_date)
        sp_use_time = _coerce_bool(sharepoint_use_time, default=download_use_time)
        sp_prefix = build_filename(sp_base, sp_use_date, sp_use_time)

        upload_result = save_excel_to_sharepoint(
            df_result,
            template_path=template_path,
            sharepoint_folder=sharepoint_folder,
            output_filename_prefix=sp_prefix,
            keep_desktop_copy=keep_desktop_copy,
            excel_visible=excel_visible,
            column_types=column_types,
        )
        sharepoint_output_path = upload_result["sharepoint_path"]
        print(f"[OK] SharePoint file saved: {sharepoint_output_path}")

    print("=" * 60)
    print("[Done] Databricks extract completed")
    if local_output_path:
        print(f"Local: {local_output_path}")
    if sharepoint_output_path:
        print(f"SharePoint: {sharepoint_output_path}")
    print("=" * 60 + "\n")

    return df_result
