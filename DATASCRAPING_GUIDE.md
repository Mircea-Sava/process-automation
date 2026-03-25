# Playwright Guide — SharePoint Automation Scripts

## Why Playwright instead of Selenium?

Selenium scripts broke when running on different laptops because:
- Selenium needs a matching `msedgedriver.exe` for the exact Edge version installed. If Edge auto-updates, the script breaks.
- CSS selectors and XPaths for SharePoint buttons varied between environments, causing "element not found" errors.

Playwright fixes both problems:
- **No driver needed** — uses your installed Edge directly (`channel="msedge"`).
- **Role-based locators** — finds buttons by their accessible role/label (e.g., `get_by_role("menuitem", name="Export")`) instead of fragile CSS classes that change between SharePoint updates.
- **Auto-waits** — waits for elements to be ready before clicking, no `time.sleep()` hacks needed.

## Setup (one-time)

```bash
pip install playwright
```

No need to run `playwright install` — our scripts use `channel="msedge"` which uses your already-installed Edge browser.

## How to create or update a script for a new SharePoint page

### Step 1: Get the SharePoint URL

Navigate to the SharePoint list/page in your browser. Copy the URL **up to and including `.aspx`** — strip off everything after the `?`.

Example:
- Full URL: `https://rtxusers.sharepoint.us/sites/.../AllItems.aspx?e=adnuBD&CID=8d67df69...`
- Use this: `https://rtxusers.sharepoint.us/sites/.../AllItems.aspx`

The query parameters (`?e=...&CID=...`) are session-specific and will differ per user. Including them can cause issues on other machines.

**Non-SharePoint sites** (e.g., `https://itdm.pwc.ca/itdm/cfm/pc_tdm_ITDM.cfm`): use the full URL as-is. The `.aspx` trimming rule only applies to SharePoint — other sites may need their query parameters to route to the correct page.

### Step 2: Run Playwright Codegen

```bash
python -m playwright codegen --channel msedge <URL>
```

Example:
```bash
python -m playwright codegen --channel msedge https://rtxusers.sharepoint.us/sites/COREPlant5MGMT-PWC2/Lists/Shipment/AllItems.aspx
```

This opens two windows:
1. **A browser** — click through the flow you want to automate (sign in, click Export, click CSV, etc.)
2. **A code panel** — records every click as Python code with exact locators

### Step 3: Copy the locators into your script

The codegen output gives you lines like:
```python
page.get_by_role("menuitem", name="Export").click()
page.get_by_role("menuitem", name="Export to CSV", exact=True).click()
```

Use these exact locators in your script. They are stable across machines because they match by **role and label**, not CSS classes.

### Key pattern for downloads

Wrap the download-triggering click with `expect_download`:
```python
with page.expect_download(timeout=60000) as download_info:
    page.get_by_role("menuitem", name="Export to CSV", exact=True).click()
download = download_info.value
csv_path = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
download.save_as(csv_path)
```

### Key pattern for waiting

Don't use `time.sleep()`. Wait for the actual element:
```python
# Wait for page to be ready — use the element you need next as the signal
page.get_by_role("menuitem", name="Export").wait_for(state="visible", timeout=120000)
```

## Troubleshooting

- **Codegen won't download browsers**: Ignore it — use `--channel msedge` to skip the bundled browser download.
- **Script works on my machine but not another**: Re-run codegen on the other machine to verify the locators match. SharePoint can render differently based on user permissions/role.
- **Element not found**: Run codegen again on that page to get updated locators.
