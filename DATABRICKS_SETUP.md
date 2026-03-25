# Databricks Setup for This Repo

This repository uses `process_automation.run_databricks_extract` from
`process_automation/databricks_extract.py` and the Databricks SQL connector.
It does not use `databricks-connect` or Spark sessions.

## 1. Install dependencies

Preferred:

```bash
pip install -e .
```

Direct dependency only:

```bash
pip install databricks-sql-connector
```

## 2. Create local `.env`

Create `process_automation/.env` (same folder as `databricks_extract.py`).
This is the default file path loaded by the code when `env_path=None`.

```env
# Required
DATABRICKS_HOST=https://<your-workspace-host>
DATABRICKS_TOKEN=<your-personal-access-token>
DATABRICKS_HTTP_PATH=/sql/1.0/warehouses/<your-warehouse-id>

# Optional fallback (only if DATABRICKS_HTTP_PATH is not set)
# DATABRICKS_ORG_ID=<workspace-org-id>
# DATABRICKS_WORKSPACE_ID=<workspace-org-id>
# DATABRICKS_CLUSTER_ID=<cluster-id>
```

Notes:
- `DATABRICKS_SERVER_HOSTNAME` is also accepted instead of `DATABRICKS_HOST`.
- `DATABRICKS_PORT` is not used by this repo's Databricks extractor.
- Preferred setup is to provide `DATABRICKS_HTTP_PATH` directly.

## 3. Keep secrets out of git

`.env` files are ignored by `.gitignore` (`*.env`), but if a token was ever
committed, rotate it immediately in Databricks.

## 4. Run the Databricks template

Use:

```bash
python templates/template_databricks_pipeline.py
```

Template file:
- `templates/template_databricks_pipeline.py`

Main options:
- `table_path` plus query-shaping (`select_columns`, `filters`, `order_by`, `limit`)
- or full `query` (overrides query-shaping)
- `save_local` and `download_dir` for local export
- `upload_to_sharepoint`, `template_path`, `sharepoint_folder` for SharePoint upload

## 5. Quick smoke test (minimal)

Set in the template:
- `upload_to_sharepoint=False`
- `save_local=True`
- `output_format="csv"`
- `limit=1`

Then run the template. If a one-row file is created, your Databricks connection
and credentials are working.

## 6. Troubleshooting (repo-specific)

- `Missing Databricks host...`
  Set `DATABRICKS_HOST` or `DATABRICKS_SERVER_HOSTNAME` in `.env`.
- `Missing Databricks token...`
  Set `DATABRICKS_TOKEN` in `.env`.
- `Missing Databricks HTTP path...`
  Set `DATABRICKS_HTTP_PATH`, or set both `DATABRICKS_ORG_ID` (or `DATABRICKS_WORKSPACE_ID`) and `DATABRICKS_CLUSTER_ID`.
- `Missing dependency 'databricks-sql-connector'`
  Install dependencies (`pip install -e .`).
- `Provide either table_path or query.`
  Set at least one of these in the template.
- `At least one output must be enabled...`
  Keep `save_local=True` or `upload_to_sharepoint=True`.

## 7. Query usage

You can either:

1. Use query-shaping with `table_path`:
   - `select_columns=["AUFNR", "WERKS"]`
   - `filters=["WERKS = '1000'"]`
   - `order_by=["AUFNR DESC"]`
   - `limit=10000`
2. Use full `query` SQL:
   - CTEs, joins, group by, window functions, unions, subqueries, etc.
   - Recommended as one `SELECT` statement for this workflow.

## 8. Connecting Power BI to Databricks

### Step 1 — Get the HTTP Path from Databricks

1. Go to https://pwc-lab.cloud.databricks.com
2. Click **Compute** in the left sidebar
3. Click on your cluster (e.g. `<your-cluster-name>`)
4. Click **Advanced Options** (at the bottom of the cluster page)
5. Click the **JDBC/ODBC** tab (not the SSH tab)
6. Copy the three values shown:
   - **Server Hostname:** `pwc-lab.cloud.databricks.com`
   - **Port:** `443`
   - **HTTP Path:** `<your-http-path>`

### Step 2 — Connect from Power BI Desktop

1. Open **Power BI Desktop**
2. Click **Get Data** (Home ribbon)
3. Search for **"Azure Databricks"** in the search box
4. Select **Azure Databricks** and click **Connect**
5. Fill in the connection details:
   - **Server Hostname:** `pwc-lab.cloud.databricks.com`
   - **HTTP Path:** `<you-http-path>`
6. Click **OK**
7. When prompted for authentication, select **Personal Access Token**
8. Paste your Databricks token (the one from your `.env` file)
9. Click **Connect**
10. Browse and select the tables you want (e.g. `team_aftermarket_mro.mro_all.vwf_sap_merged_aufk_afko_afvc_afvv`)
