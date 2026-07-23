#!/usr/bin/env python3
"""
txt_to_excel.py — Convert a plain-text file to an Excel spreadsheet.

Output is automatically named:  Battery_Alert_<Day>_<DD>_<Month>_<YYYY>.xlsx

Usage:
    python txt_to_excel.py input.txt [--delimiter DELIM] [--no-header] [--no-region]
                                      [--no-pivot] [--no-soc] [--lookup FILE]
                                      [--aws-profile PROFILE]

Arguments:
    input.txt       Path to the source text file (required)
    --delimiter     Column separator: 'tab', 'comma', 'semicolon', 'pipe',
                    or any single character (default: auto-detect)
    --no-header     Treat the first row as data, not a header
    --no-region     Skip the automatic Region column extraction
    --no-pivot      Skip the automatic Pivot sheet
    --no-soc        Skip the automatic Recent SoC lookup from Athena
    --lookup FILE   Reference workbook (e.g. IP_Scrapped_data.xlsx) whose
                    first sheet maps 'Organisation Unit' -> 'Parent
                    Organisation Unit'. Used to resolve Region first,
                    before falling back to dash-splitting.
    --aws-profile   Named AWS profile to use for the Athena query
                    (default: whatever boto3's default credential chain finds)

Region extraction:
    For each row's 'Organisation Unit' value:
      1. If --lookup was given, look it up there (exact match) and use
         the 'Parent Organisation Unit' as Region.
      2. If that misses, fall back to splitting a 'CCBSA - Bloemfontein'
         style value on ' - ' and using the part after the dash.
      3. Otherwise Region is left blank.
    The original Organisation Unit column is left unchanged.

Date extraction:
    If any column contains timestamps like "15 Jul 2026, 10:38", the script
    detects that column and appends a hidden "Date" helper column (a real
    Excel date value, not text) used by the Pivot sheet's formulas. It's
    hidden by default since it's not meant to be read directly — unhide it
    in Excel if you want to see it.

Recent SoC lookup:
    Reads the distinct values in the 'Battery' column (e.g. "P14-000295")
    and treats each as a device_serial. Runs a single live Athena query
    against analytics.batteries_v3 (workgroup data_consumers) restricted
    to the last 24 hours (configurable) that returns the most recent
    reading per device_serial within that window (using sequence_number as
    a tiebreaker for readings sharing the same time_stamp), then appends
    two columns:
      - "Recent SoC (%)"        — the latest state-of-charge reading
      - "Recent SoC Updated At" — the timestamp of that reading
    Batteries with no reading in the lookback window (or not found at
    all) are left blank in both columns. The lookback window exists so
    the query doesn't have to scan the table's entire history to find
    the most recent row per battery, which can otherwise time out on a
    large partitioned table. Adjust it with --soc-lookback-hours or
    ATHENA_SOC_LOOKBACK_HOURS if 24h is too narrow for your fleet's
    reporting cadence.
    Requires boto3 and valid AWS credentials (env vars, ~/.aws/credentials,
    or --aws-profile). Skip with --no-soc if you don't want this (e.g. no
    network access to AWS from this machine).

    KNOWN ISSUE (as of July 2026): analytics.batteries_v3 currently takes
    100+ seconds to run even trivial queries (confirmed via the Athena
    console: a single-device lookup with LIMIT 1 scanned ~30 KB but took
    over 2 minutes; a plain file-count query against the Iceberg
    "$files" metadata table timed out entirely after ~30 minutes). This
    points to the underlying table having an excessive number of small
    files that haven't been compacted, despite its
    'optimize_rewrite_data_file_threshold'='1' setting — this is a
    data-platform/compaction issue, not something fixable from the query
    side. If the Recent SoC lookup is consistently slow or timing out,
    that's very likely why. Flag it to whoever owns the batteries_v3
    pipeline (ask about running/scheduling OPTIMIZE ... EXECUTE) rather
    than re-tuning this query further. Use --no-soc to skip the lookup
    entirely in the meantime.

    Configuration is read from environment variables, which can be set via
    a .env file (loaded automatically with python-dotenv if it's installed)
    in the working directory the script is run from — the same keys used
    by the fleet-battery-report project's config.py, so both can share one
    .env. Recognised keys, all optional:
        ATHENA_DATABASE=analytics
        ATHENA_TABLE=batteries_v3
        ATHENA_WORKGROUP=data_consumers
        ATHENA_BEM_ID_COL=device_serial
        ATHENA_SOC_COL=soc
        ATHENA_TIMESTAMP_COL=time_stamp
        AWS_PROFILE=your-profile-name
        AWS_REGION=af-south-1          (or AWS_DEFAULT_REGION)
        ATHENA_S3_OUTPUT=s3://...      (omit to use the workgroup default)
        ATHENA_SOC_LOOKBACK_HOURS=24   (how far back to scan for a recent reading)
    See .env.example for a template.

Pivot sheet:
    If the Data sheet ends up with "Date" and "Type" columns, a second
    sheet named "Pivot" is added:
      - Main table: rows are each Date (+ Organisation Unit + Region, if
        present) combination in the data, columns are each distinct Alert
        Type, cells are live COUNTIFS formulas.
      - "Alerts per Day" table: one row per date with a single total
        alert count for that day, regardless of Type/Org Unit/Region.
    All values are formulas, so the pivot recalculates if Data changes.

Examples:
    python txt_to_excel.py data.txt
    python txt_to_excel.py data.csv --delimiter comma
    python txt_to_excel.py data.tsv --delimiter tab
    python txt_to_excel.py data.txt --no-soc
    python txt_to_excel.py data.txt --aws-profile balancell-analytics
"""

import sys
import os
import time
import argparse
from datetime import datetime, timedelta, timezone

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Missing dependency. Run:  pip install openpyxl")

try:
    from dotenv import load_dotenv
    load_dotenv()  # loads a .env file in the current working directory, if present
except ImportError:
    # python-dotenv is optional: if it's not installed, we just rely on
    # whatever environment variables are already set (or the hardcoded
    # defaults below). Only needed if you actually want to use a .env file.
    pass


DELIMITER_MAP = {
    "tab":       "\t",
    "comma":     ",",
    "semicolon": ";",
    "pipe":      "|",
}

REGION_SPLIT = " - "
DATETIME_FORMAT = "%d %b %Y, %H:%M"  # e.g. "15 Jul 2026, 10:38"

HEADER_FONT  = Font(name="Arial", bold=True, color="FFFFFF")
HEADER_FILL  = PatternFill("solid", start_color="2E5B9A")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
DATA_FONT    = Font(name="Arial", size=10)
BOLD_FONT    = Font(name="Arial", size=10, bold=True)
ALT_FILL     = PatternFill("solid", start_color="EEF2F7")

# ---------------------------------------------------------------------------
# Athena configuration — populated from environment variables (e.g. via a
# .env file loaded by python-dotenv above), matching the same keys used by
# the fleet-battery-report project's config.py. Every value is optional:
# region and s3_output default to None (letting boto3/the AWS profile
# decide), the rest fall back to the current hardcoded defaults if unset.
#
# .env keys recognised:
#   ATHENA_DATABASE, ATHENA_TABLE, ATHENA_WORKGROUP,
#   ATHENA_BEM_ID_COL, ATHENA_SOC_COL, ATHENA_TIMESTAMP_COL,
#   AWS_PROFILE, AWS_REGION (or AWS_DEFAULT_REGION), ATHENA_S3_OUTPUT
# ---------------------------------------------------------------------------
ATHENA_WORKGROUP      = os.getenv("ATHENA_WORKGROUP", "data_consumers")
ATHENA_DATABASE       = os.getenv("ATHENA_DATABASE", "analytics")
ATHENA_TABLE          = os.getenv("ATHENA_TABLE", "batteries_v3")
ATHENA_BEM_ID_COL     = os.getenv("ATHENA_BEM_ID_COL", "device_serial")  # matches the 'Battery' value, e.g. P14-000295
ATHENA_SOC_COL        = os.getenv("ATHENA_SOC_COL", "soc")               # state-of-charge column (assumed 0-100 numeric)
ATHENA_TIMESTAMP_COL  = os.getenv("ATHENA_TIMESTAMP_COL", "time_stamp")  # reading timestamp column
ATHENA_HOUR_INSERTED_COL = os.getenv("ATHENA_HOUR_INSERTED_COL", "hour_inserted")  # ingestion-hour column (format YYYY-MM-DD-HH), used for file-level pruning since batteries_v3 is only partitioned by device_serial
DEFAULT_AWS_PROFILE   = os.getenv("AWS_PROFILE")  # None if unset -> boto3's default credential chain
DEFAULT_AWS_REGION    = os.getenv("AWS_REGION") or os.getenv("AWS_DEFAULT_REGION")  # None -> defer to profile/config
DEFAULT_S3_OUTPUT     = os.getenv("ATHENA_S3_OUTPUT")  # None -> use the workgroup's own result location
SOC_LOOKBACK_HOURS     = float(os.getenv("ATHENA_SOC_LOOKBACK_HOURS", "24"))  # how far back to scan for a "recent" reading

BATTERY_COLUMN_NAME   = "Battery"    # column in the input .txt that holds the device_serial
SOC_POLL_INTERVAL_SEC = 2.0
SOC_POLL_TIMEOUT_SEC  = float(os.getenv("ATHENA_QUERY_TIMEOUT_SEC", "240"))


def build_output_name() -> str:
    """Return Battery_Alert_<Weekday>_<DD>_<Month>_<YYYY>.xlsx using today's date."""
    now = datetime.now()
    return now.strftime("Battery_Alert_%A_%d_%B_%Y.xlsx")


def detect_delimiter(lines: list) -> str:
    candidates = ["\t", ",", ";", "|"]
    sample = lines[:min(5, len(lines))]
    counts = {d: sum(line.count(d) for line in sample) for d in candidates}
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else None


def read_txt(path: str, delimiter) -> list:
    with open(path, encoding="utf-8", errors="replace") as f:
        lines = [l.rstrip("\n") for l in f if l.strip()]

    if not lines:
        sys.exit("The input file is empty.")

    if delimiter is None:
        delimiter = detect_delimiter(lines)

    if delimiter:
        rows = [line.split(delimiter) for line in lines]
    else:
        rows = [[line] for line in lines]

    return rows


# ---------------------------------------------------------------------------
# Region extraction
# ---------------------------------------------------------------------------

def load_region_lookup(path: str) -> dict:
    """
    Reads a reference workbook (e.g. IP_Scrapped_data.xlsx) and builds an
    {Organisation Unit: Parent Organisation Unit} dict from its FIRST
    sheet only (e.g. 'Sheet1'), which is expected to have two columns:
    'Organisation Unit' and 'Parent Organisation Unit'. The first row is
    treated as a header and skipped. If an Organisation Unit repeats, the
    first mapping seen wins.
    """
    lookup_wb = openpyxl.load_workbook(path, data_only=True)
    ws = lookup_wb.worksheets[0]

    lookup = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue  # header row
        if not row or row[0] is None:
            continue
        org_unit = str(row[0]).strip()
        parent = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        if org_unit and org_unit not in lookup:
            lookup[org_unit] = parent

    return lookup


def find_region_column(rows: list, has_header: bool) -> int:
    """
    Scan data rows to find the column with the most values containing the
    REGION_SPLIT pattern (e.g. " - "). Unlike a strict majority vote, this
    only needs at least one match, since region info is often present in
    just a subset of rows (e.g. only some Organisation Units have a region
    suffix, others are single-word names). Returns the column index with
    the highest match count, or -1 if no column has any matches.
    """
    data_rows = rows[1:] if has_header else rows
    if not data_rows:
        return -1

    num_cols = max(len(r) for r in rows)
    best_col, best_count = -1, 0

    for c in range(num_cols):
        values = [r[c] for r in data_rows if c < len(r) and r[c].strip()]
        if not values:
            continue
        matches = sum(1 for v in values if REGION_SPLIT in v)
        if matches > best_count:
            best_count = matches
            best_col = c

    return best_col


def add_region_column(rows: list, has_header: bool, region_lookup: dict = None) -> list:
    """
    Appends a 'Region' column. Resolution order per row:
      1. If a region_lookup dict is supplied, look up the row's
         'Organisation Unit' value in it (exact match).
      2. If that misses (or no lookup was supplied), fall back to
         splitting a 'CCBSA - Bloemfontein' style value on ' - '.
      3. Otherwise Region is left blank.
    The source column(s) are left unchanged.
    """
    headers = rows[0] if has_header else None

    # Which column holds the Organisation Unit / dash-style value.
    if headers and "Organisation Unit" in headers:
        org_col = headers.index("Organisation Unit")
    else:
        org_col = find_region_column(rows, has_header)

    if org_col == -1:
        return rows  # nothing to key off of

    new_rows = []
    for r_idx, row in enumerate(rows):
        row = list(row)  # copy
        if has_header and r_idx == 0:
            row.append("Region")
        else:
            value = row[org_col].strip() if org_col < len(row) and row[org_col] else ""

            region = ""
            if region_lookup is not None:
                region = region_lookup.get(value, "")
            if not region and REGION_SPLIT in value:
                region = value.split(REGION_SPLIT, 1)[1].strip()

            row.append(region)
        new_rows.append(row)

    return new_rows


# ---------------------------------------------------------------------------
# Date / Hour extraction
# ---------------------------------------------------------------------------

def parse_datetime(value: str):
    try:
        return datetime.strptime(value.strip(), DATETIME_FORMAT)
    except (ValueError, AttributeError):
        return None


def find_datetime_column(rows: list, has_header: bool) -> int:
    """Return the column index with the most parseable timestamps, or -1."""
    data_rows = rows[1:] if has_header else rows
    if not data_rows:
        return -1

    num_cols = max(len(r) for r in rows)
    best_col, best_count = -1, 0

    for c in range(num_cols):
        values = [r[c] for r in data_rows if c < len(r) and r[c].strip()]
        if not values:
            continue
        matches = sum(1 for v in values if parse_datetime(v))
        if matches > best_count:
            best_count = matches
            best_col = c

    return best_col


def add_date_hour_columns(rows: list, has_header: bool) -> list:
    """
    Detects a timestamp column like '15 Jul 2026, 10:38', keeps it
    unchanged, and appends a 'Date' (date object) column. Rows whose
    timestamp can't be parsed get a blank Date.
    """
    dt_col = find_datetime_column(rows, has_header)
    if dt_col == -1:
        return rows  # nothing to split

    new_rows = []
    for r_idx, row in enumerate(rows):
        row = list(row)
        if has_header and r_idx == 0:
            row.append("Date")
        else:
            value = row[dt_col] if dt_col < len(row) else ""
            dt = parse_datetime(value)
            row.append(dt.date() if dt else "")
        new_rows.append(row)

    return new_rows


# ---------------------------------------------------------------------------
# Recent SoC lookup (Athena)
# ---------------------------------------------------------------------------

def _run_athena_query(
    sql: str,
    aws_profile: str = None,
    region: str = None,
    s3_output: str = None,
) -> list:
    """
    Runs a query against Athena via boto3, polls until it finishes, and
    returns the results as a list of dicts (one per row, keyed by column
    name).

    Only passes region_name to the boto3 client if `region` is explicitly
    given; otherwise defers to whatever the profile/AWS config already
    specifies, to avoid region/credential mismatches. Same for s3_output:
    only sets ResultConfiguration if explicitly given, otherwise relies on
    the workgroup's own configured result location.
    """
    try:
        import boto3
    except ImportError:
        sys.exit(
            "Missing dependency for --no-soc lookup. Run:  pip install boto3\n"
            "Or pass --no-soc to skip the Athena Recent SoC lookup."
        )

    session = boto3.Session(profile_name=aws_profile) if aws_profile else boto3.Session()
    client = session.client("athena", region_name=region) if region else session.client("athena")

    start_kwargs = {
        "QueryString": sql,
        "QueryExecutionContext": {"Database": ATHENA_DATABASE},
        "WorkGroup": ATHENA_WORKGROUP,
    }
    if s3_output:
        start_kwargs["ResultConfiguration"] = {"OutputLocation": s3_output}

    start = client.start_query_execution(**start_kwargs)
    query_id = start["QueryExecutionId"]
    print(f"  Athena query started (QueryExecutionId={query_id}) — "
          f"check the Athena console 'Recent queries' tab to watch it live")

    elapsed = 0.0
    last_reported = 0
    while elapsed < SOC_POLL_TIMEOUT_SEC:
        status = client.get_query_execution(QueryExecutionId=query_id)
        state = status["QueryExecution"]["Status"]["State"]
        if state == "SUCCEEDED":
            stats = status["QueryExecution"].get("Statistics", {})
            scanned_bytes = stats.get("DataScannedInBytes")
            exec_ms = stats.get("EngineExecutionTimeInMillis")
            if scanned_bytes is not None:
                print(f"  Query succeeded — scanned {scanned_bytes / 1e6:.1f} MB "
                      f"in {(exec_ms or 0) / 1000:.1f}s")
            break
        if state in ("FAILED", "CANCELLED"):
            reason = status["QueryExecution"]["Status"].get("StateChangeReason", "unknown error")
            sys.exit(f"Athena query {state.lower()}: {reason}")
        if elapsed - last_reported >= 10:
            print(f"  ...still running ({state}), {elapsed:.0f}s elapsed")
            last_reported = elapsed
        time.sleep(SOC_POLL_INTERVAL_SEC)
        elapsed += SOC_POLL_INTERVAL_SEC
    else:
        sys.exit(
            f"Athena query timed out after {SOC_POLL_TIMEOUT_SEC:.0f}s "
            f"(QueryExecutionId={query_id}).\n"
            f"Known issue: analytics.batteries_v3 has been observed taking "
            f"100s-30min+ even for trivial single-device or metadata-only "
            f"queries, likely due to an excessive number of small files "
            f"that haven't been compacted. This is a data-platform issue, "
            f"not something fixable by tuning this query further — flag it "
            f"to whoever owns the batteries_v3 pipeline. In the meantime, "
            f"rerun with --no-soc to skip this lookup, or raise "
            f"ATHENA_QUERY_TIMEOUT_SEC in .env if you'd rather just wait "
            f"longer."
        )

    results = []
    columns = None
    paginator = client.get_paginator("get_query_results")
    for page in paginator.paginate(QueryExecutionId=query_id):
        rows = page["ResultSet"]["Rows"]
        if columns is None:
            columns = [c.get("VarCharValue", "") for c in rows[0]["Data"]]
            rows = rows[1:]  # skip header row on the first page only
        for row in rows:
            values = [cell.get("VarCharValue") for cell in row["Data"]]
            results.append(dict(zip(columns, values)))

    return results


def get_recent_soc(
    bem_ids: list,
    aws_profile: str = None,
    region: str = None,
    s3_output: str = None,
    lookback_hours: float = None,
) -> dict:
    """
    Queries analytics.batteries_v3 for the most recent SoC reading per
    bem_id in `bem_ids`, restricted to the last `lookback_hours` (default
    SOC_LOOKBACK_HOURS) to avoid scanning the entire table history.

    Matches the pattern used elsewhere in fleet_battery_report's query
    builders:
      - The time cutoff is computed here in Python and embedded as a
        literal `timestamp '...'`, NOT `current_timestamp - INTERVAL`.
        The latter is evaluated at query-execution time rather than
        planning time, which can block Athena's partition pruning
        entirely and force a full-table scan (this was the likely cause
        of the previous timeout).
      - Rows are deduplicated first via max_by(soc, time_received)
        GROUP BY device_serial, bem_id, sequence_number (same as
        missed_opportunity_charging's `deduplicated` CTE), before
        picking the most recent per device_serial.
      - An additional `hour_inserted >= '<cutoff hour>'` filter is added
        alongside the time_stamp filter. batteries_v3 is only PARTITIONED
        BY device_serial, so a time_stamp-only filter still forces a scan
        of every file in that device's entire partition history.
        hour_inserted marks the ingestion hour, and files are written (and
        very likely clustered) in ingestion order, so filtering on it lets
        Iceberg's file-level min/max stats actually skip files. This is
        safe (never excludes a row we want) because a record can't be
        received before it was generated, so time_received — and
        therefore hour_inserted — is always >= time_stamp.

    Returns {bem_id: (soc_str, timestamp_str)}. bem_ids not found within
    the lookback window are simply absent from the returned dict.
    """
    if not bem_ids:
        return {}

    if lookback_hours is None:
        lookback_hours = SOC_LOOKBACK_HOURS

    cutoff = datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(hours=lookback_hours)
    cutoff_literal = cutoff.strftime("%Y-%m-%d %H:%M:%S")
    cutoff_hour_literal = cutoff.strftime("%Y-%m-%d-%H")

    # Escape single quotes defensively, then quote each id for the IN list.
    quoted_ids = ", ".join("'" + bid.replace("'", "''") + "'" for bid in bem_ids)

    sql = f"""
        WITH deduplicated AS (
            SELECT
                {ATHENA_BEM_ID_COL},
                bem_id,
                sequence_number,
                max_by({ATHENA_SOC_COL}, time_received) AS {ATHENA_SOC_COL},
                CAST(max_by({ATHENA_TIMESTAMP_COL}, time_received) AS timestamp) AS {ATHENA_TIMESTAMP_COL}
            FROM {ATHENA_DATABASE}.{ATHENA_TABLE}
            WHERE {ATHENA_BEM_ID_COL} IN ({quoted_ids})
              AND {ATHENA_HOUR_INSERTED_COL} >= '{cutoff_hour_literal}'
              AND {ATHENA_TIMESTAMP_COL} >= timestamp '{cutoff_literal}'
            GROUP BY {ATHENA_BEM_ID_COL}, bem_id, sequence_number
        ),
        ranked AS (
            SELECT *,
                ROW_NUMBER() OVER (
                    PARTITION BY {ATHENA_BEM_ID_COL}
                    ORDER BY {ATHENA_TIMESTAMP_COL} DESC, sequence_number DESC
                ) AS rn
            FROM deduplicated
        )
        SELECT {ATHENA_BEM_ID_COL}, {ATHENA_SOC_COL}, {ATHENA_TIMESTAMP_COL}
        FROM ranked
        WHERE rn = 1
    """

    rows = _run_athena_query(sql, aws_profile=aws_profile, region=region, s3_output=s3_output)

    soc_by_id = {}
    for r in rows:
        bem_id = r.get(ATHENA_BEM_ID_COL)
        if bem_id:
            soc_by_id[bem_id] = (r.get(ATHENA_SOC_COL), r.get(ATHENA_TIMESTAMP_COL))

    return soc_by_id


def add_recent_soc_column(
    rows: list,
    has_header: bool,
    aws_profile: str = None,
    region: str = None,
    s3_output: str = None,
    lookback_hours: float = None,
) -> list:
    """
    Appends 'Recent SoC (%)' and 'Recent SoC Updated At' columns, keyed
    on the 'Battery' column (treated as bem_id). Batteries not found in
    Athena within the lookback window, or blank/placeholder ('-') values,
    are left blank.
    """
    if not has_header:
        print("  (Skipping Recent SoC: --no-header means there's no 'Battery' column to key off)")
        return rows

    headers = rows[0]
    if BATTERY_COLUMN_NAME not in headers:
        print(f"  (Skipping Recent SoC: no '{BATTERY_COLUMN_NAME}' column found)")
        return rows

    battery_col = headers.index(BATTERY_COLUMN_NAME)
    data_rows = rows[1:]

    bem_ids = sorted({
        row[battery_col].strip()
        for row in data_rows
        if battery_col < len(row) and row[battery_col].strip() and row[battery_col].strip() != "-"
    })

    effective_lookback = SOC_LOOKBACK_HOURS if lookback_hours is None else lookback_hours
    print(f"  Recent SoC: looking up {len(bem_ids)} batteries from Athena "
          f"(last {effective_lookback:g}h) — note: analytics.batteries_v3 has "
          f"a known compaction/small-files issue that can make even trivial "
          f"queries slow; this may take a while")
    soc_by_id = get_recent_soc(
        bem_ids, aws_profile=aws_profile, region=region, s3_output=s3_output,
        lookback_hours=effective_lookback,
    )
    matched = sum(1 for bid in bem_ids if bid in soc_by_id)
    print(f"  Recent SoC: matched {matched}/{len(bem_ids)} batteries in Athena "
          f"(last {effective_lookback:g}h)")

    new_rows = []
    for r_idx, row in enumerate(rows):
        row = list(row)
        if r_idx == 0:
            row.append("Recent SoC (%)")
            row.append("Recent SoC Updated At")
        else:
            value = row[battery_col].strip() if battery_col < len(row) and row[battery_col] else ""
            soc, ts = soc_by_id.get(value, (None, None))
            row.append(soc if soc is not None else "")
            row.append(ts if ts is not None else "")
        new_rows.append(row)

    return new_rows


# ---------------------------------------------------------------------------
# Excel writing
# ---------------------------------------------------------------------------

def auto_width(sheet):
    for col in sheet.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except Exception:
                pass
        sheet.column_dimensions[col_letter].width = min(max_len + 4, 60)


def write_data_sheet(wb, rows: list, has_header: bool = True):
    ws = wb.active
    ws.title = "Data"
    headers = rows[0] if has_header else None

    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            is_header_row = has_header and r_idx == 1
            header_name = headers[c_idx - 1] if headers and c_idx - 1 < len(headers) else None

            if is_header_row:
                cell = ws.cell(row=r_idx, column=c_idx, value=str(value).strip())
                cell.font = HEADER_FONT
                cell.fill = HEADER_FILL
                cell.alignment = HEADER_ALIGN
                continue

            if header_name == "Date" and value != "":
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.number_format = "dd mmm yyyy"
            elif header_name == "Hour" and value != "":
                cell = ws.cell(row=r_idx, column=c_idx, value=int(value))
                cell.number_format = "0"
            elif header_name == "Recent SoC (%)" and value not in ("", None):
                try:
                    numeric_soc = float(value)
                    cell = ws.cell(row=r_idx, column=c_idx, value=numeric_soc)
                    cell.number_format = "0.0"
                except (TypeError, ValueError):
                    cell = ws.cell(row=r_idx, column=c_idx, value=str(value).strip())
            else:
                cell = ws.cell(
                    row=r_idx, column=c_idx,
                    value=str(value).strip() if value != "" else None
                )

            cell.font = DATA_FONT
            if r_idx % 2 == 0:
                cell.fill = ALT_FILL

    if has_header:
        ws.freeze_panes = "A2"

    auto_width(ws)

    # The 'Date' column is a helper for the Pivot sheet's formulas, not
    # something the user asked to see on the main sheet — keep it but hide it.
    if headers and "Date" in headers:
        date_col_idx = headers.index("Date") + 1
        ws.column_dimensions[get_column_letter(date_col_idx)].hidden = True

    return ws


def add_pivot_sheet(wb, rows: list, has_header: bool):
    """
    Adds a 'Pivot' sheet summarizing alert counts by Date, Organisation
    Unit, and Region (rows) and Type (columns), using live COUNTIFS
    formulas against the Data sheet. Requires 'Date' and 'Type' columns;
    'Organisation Unit' and 'Region' are included automatically if
    present, but are not required. Also adds an 'Alerts per Day' summary
    table with one row per date and a single total-alerts count.
    """
    if not has_header:
        return None

    headers = rows[0]
    required = ("Date", "Type")
    if not all(h in headers for h in required):
        return None

    # Row-grouping fields: Date always, plus Organisation Unit and
    # Region when the Data sheet has them.
    group_fields = ["Date"]
    for optional in ("Organisation Unit", "Region"):
        if optional in headers:
            group_fields.append(optional)

    type_idx = headers.index("Type")
    group_idxs = [headers.index(f) for f in group_fields]

    data_rows = rows[1:]
    n_data_rows = len(rows)  # includes header row; last data row = this number

    def group_key(r):
        return tuple(
            r[i] if i < len(r) else "" for i in group_idxs
        )

    combos = sorted({
        group_key(r) for r in data_rows
        if (r[group_idxs[0]] if group_idxs[0] < len(r) else "") != ""
    }, key=lambda combo: tuple(str(v) for v in combo))
    types = sorted({
        r[type_idx] for r in data_rows
        if type_idx < len(r) and str(r[type_idx]).strip()
    })

    if not combos or not types:
        return None

    group_col_letters = [get_column_letter(i + 1) for i in group_idxs]
    type_col_letter = get_column_letter(type_idx + 1)

    ws = wb.create_sheet("Pivot")
    n_groups = len(group_fields)

    # Header row
    for c_offset, field in enumerate(group_fields):
        ws.cell(row=1, column=1 + c_offset, value=field)
    for c_offset, t in enumerate(types):
        ws.cell(row=1, column=1 + n_groups + c_offset, value=t)
    total_col = 1 + n_groups + len(types)
    ws.cell(row=1, column=total_col, value="Total")

    for c_idx in range(1, total_col + 1):
        cell = ws.cell(row=1, column=c_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN

    # Data rows: one per group-field combination present in the source
    for r_offset, combo in enumerate(combos):
        r_idx = 2 + r_offset

        for c_offset, (field, value) in enumerate(zip(group_fields, combo)):
            c_idx = 1 + c_offset
            cell = ws.cell(row=r_idx, column=c_idx, value=value if value != "" else None)
            cell.font = DATA_FONT
            if field == "Date":
                cell.number_format = "dd mmm yyyy"
            elif field == "Hour":
                cell.number_format = "0"

        for c_offset, t in enumerate(types):
            c_idx = 1 + n_groups + c_offset
            type_header_ref = f"{get_column_letter(c_idx)}$1"
            criteria = []
            for g_offset, col_letter in enumerate(group_col_letters):
                group_value = combo[g_offset]
                if group_value in (None, ""):
                    criterion_expr = '""'
                else:
                    criterion_expr = f"${get_column_letter(1 + g_offset)}{r_idx}"
                criteria.append(
                    f"Data!${col_letter}$2:${col_letter}${n_data_rows},{criterion_expr}"
                )
            criteria.append(
                f"Data!${type_col_letter}$2:${type_col_letter}${n_data_rows},{type_header_ref}"
            )
            formula = "=COUNTIFS(" + ",".join(criteria) + ")"
            cell = ws.cell(row=r_idx, column=c_idx, value=formula)
            cell.font = DATA_FONT

        first_type_letter = get_column_letter(1 + n_groups)
        last_type_letter = get_column_letter(n_groups + len(types))
        total_cell = ws.cell(
            row=r_idx, column=total_col,
            value=f"=SUM({first_type_letter}{r_idx}:{last_type_letter}{r_idx})"
        )
        total_cell.font = BOLD_FONT

        if r_idx % 2 == 0:
            for c_idx in range(1, total_col + 1):
                ws.cell(row=r_idx, column=c_idx).fill = ALT_FILL

    # Grand total row
    last_data_row = 1 + len(combos)
    grand_row = last_data_row + 1
    ws.cell(row=grand_row, column=1, value="Grand Total").font = BOLD_FONT
    for c_idx in range(1 + n_groups, total_col + 1):
        col_letter = get_column_letter(c_idx)
        cell = ws.cell(
            row=grand_row, column=c_idx,
            value=f"=SUM({col_letter}2:{col_letter}{last_data_row})"
        )
        cell.font = BOLD_FONT

    # -----------------------------------------------------------------
    # "Alerts per Day" summary table: one row per date, single total
    # count of alerts that day (regardless of Type / Org Unit / Region).
    # -----------------------------------------------------------------
    date_idx = headers.index("Date")
    date_col_letter = get_column_letter(date_idx + 1)

    dates = sorted({
        r[date_idx] for r in data_rows
        if date_idx < len(r) and r[date_idx] != ""
    })

    summary_start_row = grand_row + 3
    ws.cell(row=summary_start_row, column=1, value="Alerts per Day").font = BOLD_FONT

    header_row = summary_start_row + 1
    ws.cell(row=header_row, column=1, value="Date")
    ws.cell(row=header_row, column=2, value="Total Alerts")
    for c_idx in (1, 2):
        cell = ws.cell(row=header_row, column=c_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN

    for r_offset, date_val in enumerate(dates):
        r_idx = header_row + 1 + r_offset
        date_cell = ws.cell(row=r_idx, column=1, value=date_val)
        date_cell.number_format = "dd mmm yyyy"
        date_cell.font = DATA_FONT

        total_cell = ws.cell(
            row=r_idx, column=2,
            value=(
                f"=COUNTIF(Data!${date_col_letter}$2:${date_col_letter}${n_data_rows},"
                f"$A{r_idx})"
            )
        )
        total_cell.font = DATA_FONT

        if r_idx % 2 == 0:
            ws.cell(row=r_idx, column=1).fill = ALT_FILL
            ws.cell(row=r_idx, column=2).fill = ALT_FILL

    summary_last_row = header_row + len(dates)
    summary_grand_row = summary_last_row + 1
    ws.cell(row=summary_grand_row, column=1, value="Grand Total").font = BOLD_FONT
    ws.cell(
        row=summary_grand_row, column=2,
        value=f"=SUM(B{header_row + 1}:B{summary_last_row})"
    ).font = BOLD_FONT

    ws.freeze_panes = "A2"
    auto_width(ws)
    return ws


def main():
    parser = argparse.ArgumentParser(description="Convert a .txt file to Excel (.xlsx)")
    parser.add_argument("input", help="Source text file")
    parser.add_argument("--delimiter", default=None,
                        help="Column delimiter: tab, comma, semicolon, pipe, or a character")
    parser.add_argument("--no-header", action="store_true",
                        help="Treat the first row as data, not a header")
    parser.add_argument("--no-region", action="store_true",
                        help="Skip automatic Region column extraction")
    parser.add_argument("--lookup", default=None,
                        help="Path to a reference workbook (e.g. IP_Scrapped_data.xlsx) "
                             "whose first sheet maps 'Organisation Unit' -> "
                             "'Parent Organisation Unit'. Used to resolve Region "
                             "before falling back to dash-splitting.")
    parser.add_argument("--no-pivot", action="store_true",
                        help="Skip automatic Pivot sheet")
    parser.add_argument("--no-soc", action="store_true",
                        help="Skip the automatic Recent SoC lookup from Athena")
    parser.add_argument("--aws-profile", default=DEFAULT_AWS_PROFILE,
                        help="Named AWS profile to use for the Athena query "
                             "[env: AWS_PROFILE] (default: boto3's default "
                             "credential chain if unset)")
    parser.add_argument("--region", default=DEFAULT_AWS_REGION,
                        help="AWS region for the Athena client "
                             "[env: AWS_REGION or AWS_DEFAULT_REGION] "
                             "(default: deferred to the profile/AWS config)")
    parser.add_argument("--s3-output", default=DEFAULT_S3_OUTPUT,
                        help="S3 URI for Athena query results "
                             "[env: ATHENA_S3_OUTPUT] "
                             "(default: the workgroup's own result location)")
    parser.add_argument("--soc-lookback-hours", type=float, default=SOC_LOOKBACK_HOURS,
                        help="How many hours back to scan for a 'recent' SoC "
                             "reading [env: ATHENA_SOC_LOOKBACK_HOURS] (default: 24)")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        sys.exit(f"File not found: {args.input}")

    # Auto-generate output filename
    output_path = build_output_name()

    # Place the output in the same directory as the input file
    input_dir = os.path.dirname(os.path.abspath(args.input))
    output_path = os.path.join(input_dir, output_path)

    # Resolve delimiter
    delim = args.delimiter
    if delim and delim.lower() in DELIMITER_MAP:
        delim = DELIMITER_MAP[delim.lower()]

    rows = read_txt(args.input, delim)
    has_header = not args.no_header

    region_lookup = None
    if args.lookup:
        if not os.path.isfile(args.lookup):
            sys.exit(f"Lookup file not found: {args.lookup}")
        region_lookup = load_region_lookup(args.lookup)

    if not args.no_region:
        rows = add_region_column(rows, has_header, region_lookup=region_lookup)

    if not args.no_pivot:
        rows = add_date_hour_columns(rows, has_header)

    if not args.no_soc:
        rows = add_recent_soc_column(
            rows, has_header,
            aws_profile=args.aws_profile,
            region=args.region,
            s3_output=args.s3_output,
            lookback_hours=args.soc_lookback_hours,
        )

    wb = openpyxl.Workbook()
    write_data_sheet(wb, rows, has_header=has_header)

    pivot_ws = None
    if not args.no_pivot:
        pivot_ws = add_pivot_sheet(wb, rows, has_header)

    wb.save(output_path)

    print(f"Done! {len(rows)} rows written to:")
    print(f"  {output_path}")
    if pivot_ws is not None:
        print("  + Pivot sheet (Date x Alert Type, plus Alerts per Day)")
    elif not args.no_pivot:
        print("  (No Pivot sheet: couldn't find Date/Type columns)")


if __name__ == "__main__":
    main()