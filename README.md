# Battery Alerts

Battery Alerts turns a tabular text export of open battery alerts into
formatted Excel reports, grouped by region. It can enrich reports with
battery state-of-charge (SoC) telemetry from Snowflake and optional region
and customer-reference information from Odoo.

## What it does

For each input alert row, the report can:

- Add the SoC at the time the alert started and the battery's latest SoC.
- Mark batteries with no recent telemetry as offline or show when they last
  communicated.
- Resolve a region from an organisation-unit workbook, the organisation-unit
  name, or (when configured) Odoo.
- Add date/pivot information, alert age, and a stuck-low-SoC flag.
- Write one Excel workbook per region and a ZIP containing the dated report
  folder.

The main entry point is `BatteryAlerts.py`. Older/standalone scripts in
`archive/` are not needed for this report workflow.

## Requirements

- Python 3.10 or newer
- Access to the Snowflake account and telemetry table for SoC enrichment
- Optional: Odoo API credentials for the region fallback

## Setup

From the project directory, create and activate a virtual environment and
install the packages used by the report:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install openpyxl python-dotenv snowflake-connector-python
```

Copy `.env.example` to `.env` and set your Snowflake account, user, warehouse,
and (if needed) role. The report authenticates with Snowflake's
`externalbrowser` authenticator, which opens a browser sign-in when a query
runs. SoC data is read from
`BALANCELL_DATA_PRODUCTION.TELEMETRY.TELEMETRY_MODULE`.

```bash
cp .env.example .env
```

Edit `.env` with your own credentials and settings. Do not commit `.env` or
put passwords/API keys in the example file. The report automatically loads
`.env` from the current working directory.

Odoo settings (`ODOO_URL`, `ODOO_DB`, `ODOO_USERNAME`, and `ODOO_API_KEY`) are
optional. Without them, the Odoo region fallback is skipped. `--lookup` can
be used to provide an organisation-unit mapping workbook instead.

## Run

From the project directory, pass a tab-, comma-, semicolon-, or
pipe-delimited text export:

```bash
python BatteryAlerts.py alerts.txt
```

Delimiter detection is automatic. To specify it explicitly or use the
included reference workbook:

```bash
python BatteryAlerts.py alerts.txt --delimiter tab --lookup IP_Scrapped_data.xlsx
```

By default, the script attempts both Snowflake SoC lookups. To create a
report without Snowflake access:

```bash
python BatteryAlerts.py alerts.txt --no-soc --no-soc-at-start
```

See all available options with:

```bash
python BatteryAlerts.py --help
```

The input should normally include a header row. The `Battery` column is used
to match telemetry module serials; `Started At` is needed for SoC-at-start
and alert-age features. Region enrichment uses `Organisation Unit` when
available.

## Output

Each run writes regional Excel workbooks to a dated folder under
`data/results/`, then creates a ZIP of that folder in the same directory.
For example:

```text
data/results/
  Battery_Alert_<Weekday>_<DD>_<Month>_<YYYY>/
    Battery_Alert_<Weekday>_<DD>_<Month>_<YYYY>_<Region>.xlsx
  Battery_Alert_<Weekday>_<DD>_<Month>_<YYYY>.zip
```

Use `--no-region` to produce a single ungrouped workbook instead.
