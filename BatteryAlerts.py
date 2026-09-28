#!/usr/bin/env python3
"""
txt_to_excel.py — Convert a plain-text file to an Excel spreadsheet.

Output is automatically named:  Battery_Alert_<Day>_<DD>_<Month>_<YYYY>.xlsx

Usage:
    python txt_to_excel.py input.txt [--delimiter DELIM] [--no-header] [--no-region]
                                      [--no-pivot] [--no-soc] [--no-soc-at-start]
                                      [--soc-at-start-window-hours HOURS]
                                      [--no-ageing] [--historical-age-days DAYS]
                                      [--no-stuck-soc] [--stuck-soc-threshold-pct PCT]
                                      [--stuck-soc-min-days DAYS]
                                      [--lookup FILE] [--aws-profile PROFILE]

Arguments:
    input.txt       Path to the source text file (required)
    --delimiter     Column separator: 'tab', 'comma', 'semicolon', 'pipe',
                    or any single character (default: auto-detect)
    --no-header     Treat the first row as data, not a header
    --no-region     Skip the automatic Region column extraction
    --no-odoo-region  Skip the Odoo fallback lookup for batteries with no
                    Organisation Unit from the alert export
    --no-pivot      Skip the automatic Pivot sheet
    --no-soc        Skip the automatic Recent SoC lookup from Athena
    --no-soc-at-start  Skip the SoC-at-alert-Started-At lookup
    --soc-at-start-window-hours  How far before each alert's Started At
                    to search for a reading (default: 2)
    --no-ageing     Skip the 'Days Since Started' column and the
                    historical-overdue sort/highlight
    --historical-age-days  Alerts at least this many days old are flagged
                    'Historical — overdue' and sorted to the top (default: 75)
    --no-stuck-soc  Skip the 'SoC Health' stuck-battery flag
    --stuck-soc-threshold-pct  Recent SoC at/below this percent counts as
                    'low' (default: 5)
    --stuck-soc-min-days  Alert must be open at least this many days,
                    still at/below the SoC threshold, to be flagged as
                    stuck (default: 14)
    --lookup FILE   Reference workbook (e.g. IP_Scrapped_data.xlsx) whose
                    first sheet maps 'Organisation Unit' -> 'Parent
                    Organisation Unit'. Used to resolve Region first,
                    before falling back to dash-splitting.
    --aws-profile   Named AWS profile to use for the Athena query
                    (default: whatever boto3's default credential chain finds)

Region extraction:
    Appends 'Region' and 'Sold-To Customer'. Region's resolution order,
    per row's 'Organisation Unit' value:
      1. If --lookup was given, look it up there (exact match) and use
         the 'Parent Organisation Unit' as Region.
      2. If that misses, fall back to splitting a 'CCBSA - Bloemfontein'
         style value on ' - ' and using the part after the dash.
      3. If Organisation Unit is itself blank/'-' (so (1) and (2) had
         nothing to work with), look up the row's 'Battery' serial in
         Odoo instead — Odoo is the live system of record for where a
         battery physically is, unlike the --lookup workbook, which is a
         point-in-time export that goes stale as Organisation Units get
         created/renamed. See get_odoo_region_fallback for the exact
         query (stock.lot -> stock.quant/stock.move.line/stock.picking,
         same pattern as the fleet-battery-report Warehousing project,
         resolved to the actual customer company via
         commercial_partner_id — not the raw delivery-contact name). The
         resolved customer name is used ONLY internally here, to dash-
         split it the same way as (2) IF it has a dash — it's not
         surfaced as its own column (a company name like "EIE Industrial
         power a division of CFAO Equipment SA" repeats identically
         across every battery under that reseller, so 'Sold-To Customer'
         below — the actual per-shipment reference — carries the useful
         signal instead). Requires ODOO_URL/ODOO_DB/ODOO_USERNAME/
         ODOO_API_KEY (env vars, e.g. via .env) — silently skipped if
         unset. Skip explicitly with --no-odoo-region.
      4. Otherwise Region is left blank.
    'Sold-To Customer' is a separate field entirely from Region/Odoo's
    customer name above — the Sale Order's own client_order_ref (Odoo's
    own "Customer Reference" field, e.g. "KARGO NATIONAL
    ORD-27264-27265- PE"), a purchase-order-style reference the CUSTOMER
    supplied, run through _trim_po_prefix and populated whenever the
    battery shipped via a sale order with one on file. Despite the
    column name, it's NOT the customer's name and NOT our own SO number
    — don't confuse the two when reading this column.
    The original Organisation Unit column is left unchanged — Odoo only
    fills in Region and Sold-To Customer, it never overwrites what the
    alert export reported.

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
    Batteries with no reading in the lookback window trigger a second,
    unbounded "last-ever-seen" lookup so both columns can say which case
    applies instead of just being left blank:
      - "Battery Offline/Disconnected" if the battery has no history at
        all in the table.
      - "No Communication Since <timestamp>" if it reported before, using
        its last-known reading's timestamp, however old.
    The lookback window on the main query exists so it doesn't have to
    scan the table's entire history to find the most recent row per
    battery, which can otherwise time out on a large partitioned table.
    Adjust it with --soc-lookback-hours or ATHENA_SOC_LOOKBACK_HOURS if
    24h is too narrow for your fleet's reporting cadence.
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

SoC at Start lookup:
    Separate from Recent SoC above. For each alert row, looks up the SoC
    reading in effect at (or just before) THAT row's own 'Started At'
    timestamp — not the battery's latest reading, but the reading closest
    to when that specific alert began. Adds one column, inserted
    immediately after 'Started At':
      - "SoC at Start (%)"
    Requires both 'Battery' and 'Started At' columns. Searches up to
    ATHENA_SOC_AT_START_WINDOW_HOURS (default 2h, --soc-at-start-window-hours)
    before each alert's timestamp for the nearest reading at or before it.
    Rows with a blank/'-' Battery or unparseable Started At are left
    blank. Rows with a valid Battery/Started At but no reading found in
    that window instead get the same "Battery Offline/Disconnected" /
    "No Communication Since <timestamp>" placeholder described above for
    Recent SoC, resolved from the same unbounded last-ever-seen lookup.
    Skip with --no-soc-at-start.
    This runs as a SEPARATE Athena query from Recent SoC (matching up to
    one row per unique alert timestamp instead of one row per battery),
    so it adds its own runtime on top of whatever Recent SoC already
    takes — and is subject to the same known batteries_v3 compaction
    slowness described above.

Ageing / historical-overdue flag:
    Appends "Days Since Started" — integer day count from Started At to
    today, based on each row's 'Started At'. An alert counts as
    historical/overdue once it's at least --historical-age-days (default
    75, env HISTORICAL_AGE_DAYS) days old. This is a ROLLING threshold
    measured from today, not a fixed calendar date, so it keeps meaning
    the same thing every time this report runs, rather than needing the
    cutoff hand-updated as time passes.
    Rows are then re-sorted (stably, so each group keeps its existing
    relative order) so every historical/overdue row is pulled to the top
    of the Data sheet instead of sitting buried at the bottom of an
    otherwise reverse-chronological list, and given a highlight fill
    across the whole row for visibility — there's no separate text flag
    column for this; write_data_sheet re-derives it straight from "Days
    Since Started" against the same threshold. Requires a 'Started At'
    column. Skip with --no-ageing.

Stuck-SoC flag:
    Appends a "SoC Health" column, keyed on 'Battery', 'Started At', and
    'Recent SoC (%)': flags "Stuck at <n>% — No Recovery (<d>d)" once an
    alert's Recent SoC is at/below --stuck-soc-threshold-pct (default 5%)
    AND it's been open at least --stuck-soc-min-days (default 14) days.
    This is a heuristic from today's two numbers, not a scan of the
    battery's full history — it can't distinguish "flat the whole time"
    from "recovered partway and dropped again". Requires 'Recent SoC (%)'
    (i.e. --no-soc must not be set). Skip with --no-stuck-soc.

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
import re
import time
import argparse
import xmlrpc.client
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
NO_DATA_FONT = Font(name="Arial", size=10, italic=True, color="999999")
ALT_FILL     = PatternFill("solid", start_color="EEF2F7")
NO_DATA_TEXT = "No data provided"  # fallback for rows with no valid Battery id to check at all
BATTERY_OFFLINE_TEXT = "Battery Offline/Disconnected"  # battery has never reported a single reading
NO_COMM_PREFIX = "No Communication Since "  # battery reported before, then went quiet; suffixed with its last-known timestamp

HISTORICAL_FILL  = PatternFill("solid", start_color="FBDCD8")
# Rolling threshold, not a fixed calendar date, so "overdue" keeps meaning
# the same thing as this report keeps running week after week.
HISTORICAL_AGE_DAYS = float(os.getenv("HISTORICAL_AGE_DAYS", "75"))

STUCK_SOC_FONT = Font(name="Arial", size=10, bold=True, color="B45309")
# SoC at/below this % counts as "low"; the alert must also have been open
# at least this many days with the reading still stuck there before it's
# flagged as a likely-dead battery rather than a normal in-progress alert.
STUCK_SOC_THRESHOLD_PCT = float(os.getenv("STUCK_SOC_THRESHOLD_PCT", "5"))
STUCK_SOC_MIN_DAYS      = float(os.getenv("STUCK_SOC_MIN_DAYS", "14"))

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
SOC_AT_START_WINDOW_HOURS = float(os.getenv("ATHENA_SOC_AT_START_WINDOW_HOURS", "2"))  # how far before each alert's Started At to search for a reading

# ---------------------------------------------------------------------------
# Odoo configuration — same env vars/pattern as the fleet-battery-report
# Warehousing project's battery_warranty_etl.py, so a .env already set up
# for that project can be reused here as-is. All optional: if any are
# unset, the Odoo Region fallback below silently no-ops rather than erroring,
# so this stays inert until someone deliberately configures it.
# ---------------------------------------------------------------------------
ODOO_URL      = os.getenv("ODOO_URL")
ODOO_DB       = os.getenv("ODOO_DB")
ODOO_USERNAME = os.getenv("ODOO_USERNAME")
ODOO_API_KEY  = os.getenv("ODOO_API_KEY")
ODOO_BATCH_SIZE = 200

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


def _is_blank_or_dash(value) -> bool:
    return value is None or str(value).strip() in ("", "-")


# Matches a leading "PO" / "PO:" followed by a digit-led run of digits and
# separators (dashes, spaces, '&', commas) — e.g. "PO: 7814-7815 " or
# "PO234749 ". Deliberately anchored to the START of the string only: PO
# numbers embedded mid-string ("TOYOTA STR POLOKANE PO-28262 JHB") or at
# the end ("TRANSPORT COSTS DBN- PO7577") are left alone rather than risk
# a wrong match. client_order_ref is free text with no fixed format (a
# broad real-data sample turned up plain names, quote references, repair-
# order refs, and PO numbers in every position), so this only ever
# handles the "PO number as a clean prefix" case — most PO-bearing values
# won't match at all, and that's intentional: better to leave a PO number
# in place than to guess wrong and mangle real text elsewhere in the
# string.
_PO_PREFIX_RE = re.compile(r"^PO:?\s*\d[\d\-\s&,]*")


def _trim_po_prefix(text: str) -> str:
    """
    Strips a clean leading PO-number prefix (see _PO_PREFIX_RE) off
    `text`, e.g. "PO: 7814-7815 TOYOTA NEW SALES - PE" -> "TOYOTA NEW
    SALES - PE". If there's no match, or the match consumes the ENTIRE
    string (e.g. "PO100213" is nothing but a PO number), returns `text`
    unchanged — trimming to an empty string would throw away the only
    information the field had.
    """
    if not text:
        return text
    match = _PO_PREFIX_RE.match(text)
    if not match:
        return text
    remainder = text[match.end():].strip(" -:")
    return remainder or text


def add_region_column(
    rows: list,
    has_header: bool,
    region_lookup: dict = None,
    odoo_fallback: dict = None,
) -> list:
    """
    Appends 'Region' and 'Sold-To Customer' columns. Region's resolution
    order per row:
      1. If a region_lookup dict is supplied, look up the row's
         'Organisation Unit' value in it (exact match).
      2. If that misses (or no lookup was supplied), fall back to
         splitting a 'CCBSA - Bloemfontein' style value on ' - '.
      3. If Organisation Unit is itself blank/'-' (so (1) and (2) had
         nothing to work with) and an odoo_fallback dict is supplied,
         look up the row's 'Battery' serial in it — see
         get_odoo_region_fallback for how that's built — and dash-split
         the resolved name the same way as (2), IF it has a dash.
      4. Otherwise Region is left blank.

    Odoo's resolved name (odoo_fallback[battery]["name"]) is used ONLY
    for that dash-split — it's often a full legal customer name (e.g.
    "EIE Industrial power a division of CFAO Equipment SA") that's
    identical across every battery under that reseller, so it isn't
    surfaced as its own column; 'Sold-To Customer' below carries the
    per-shipment signal instead.

    'Sold-To Customer' is odoo_fallback's separate customer_reference
    value (the Sale Order's own client_order_ref, NOT our SO number and
    despite "Customer" in its name, NOT the customer's name either — see
    get_odoo_region_fallback), populated whenever the row's Battery has
    one on file, independent of whether Region got filled.

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

    battery_col = headers.index(BATTERY_COLUMN_NAME) if headers and BATTERY_COLUMN_NAME in headers else -1

    new_rows = []
    for r_idx, row in enumerate(rows):
        row = list(row)  # copy
        if has_header and r_idx == 0:
            row.append("Region")
            row.append("Sold-To Customer")
        else:
            value = row[org_col].strip() if org_col < len(row) and row[org_col] else ""

            region = ""
            if region_lookup is not None:
                region = region_lookup.get(value, "")
            if not region and REGION_SPLIT in value:
                region = value.split(REGION_SPLIT, 1)[1].strip()

            odoo_name = None
            customer_reference = None
            if _is_blank_or_dash(value) and odoo_fallback and battery_col != -1:
                battery = row[battery_col].strip() if battery_col < len(row) and row[battery_col] else ""
                odoo_match = odoo_fallback.get(battery) if battery else None
                if odoo_match:
                    odoo_name = odoo_match.get("name")
                    customer_reference = odoo_match.get("customer_reference")
                if not region and odoo_name and REGION_SPLIT in odoo_name:
                    region = odoo_name.split(REGION_SPLIT, 1)[1].strip()

            row.append(region)
            row.append(customer_reference or "")
        new_rows.append(row)

    return new_rows


def get_odoo_region_fallback(battery_serials: list) -> dict:
    """
    For batteries whose Organisation Unit came back blank/'-' from the
    alert export (so --lookup/dash-splitting in add_region_column had
    nothing to work with), asks Odoo — the actual system of record for
    where a battery physically is — for its current customer or location
    name, to use as an Organisation Unit stand-in for Region resolution.

    Ports the connection/fetch pattern from the fleet-battery-report
    Warehousing project's battery_warranty_etl.py (stock.lot ->
    stock.quant/stock.location for current location, stock.move.line/
    stock.picking for the customer it shipped to), scoped down to just
    the serials that actually need it instead of the whole fleet.

    Returns {battery_serial: {"name": ..., "customer_reference": ...}}.
    "name" is the actual customer company name (Odoo's
    commercial_partner_id, not the raw delivery contact — see the note
    below) if the battery is currently at a customer site, otherwise the
    stock location's complete_name (e.g. a warehouse). "customer_reference"
    is the Sale Order's client_order_ref (Odoo's own "Customer Reference"
    field, e.g. "KARGO NATIONAL ORD-27264-27265- PE" — a purchase-order-
    style reference the CUSTOMER supplied, distinct from both our own SO
    number and the customer's name) when the battery shipped via a sale
    order; None when it didn't (e.g. sitting in an internal warehouse) or
    the order has no reference on file. Run through _trim_po_prefix
    first, which strips a leading "PO: <number>" prefix when it's a
    clean match — see that function for why this is a best-effort trim,
    not a guaranteed-correct one (client_order_ref is inconsistent free
    text with no fixed format).

    Serials not found in Odoo at all, or found but sitting at a location
    with neither a name nor a customer, are absent from the result.

    Requires ODOO_URL, ODOO_DB, ODOO_USERNAME, ODOO_API_KEY (env vars,
    e.g. via .env, same keys as the Warehousing project) — returns {}
    without querying if any are unset, so this stays a no-op until Odoo
    is actually configured for this project.
    """
    if not battery_serials:
        return {}

    if not all([ODOO_URL, ODOO_DB, ODOO_USERNAME, ODOO_API_KEY]):
        return {}

    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(ODOO_DB, ODOO_USERNAME, ODOO_API_KEY, {})
    if not uid:
        print("  (Odoo Region fallback: authentication failed — check ODOO_DB/ODOO_USERNAME/ODOO_API_KEY)")
        return {}
    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")

    def search_read_all(model, domain, fields, order="id asc"):
        all_rows = []
        offset = 0
        while True:
            batch = models.execute_kw(
                ODOO_DB, uid, ODOO_API_KEY,
                model, "search_read",
                [domain],
                {"fields": fields, "limit": ODOO_BATCH_SIZE, "offset": offset, "order": order},
            )
            if not batch:
                break
            all_rows.extend(batch)
            if len(batch) < ODOO_BATCH_SIZE:
                break
            offset += len(batch)
        return all_rows

    lots = search_read_all(
        "stock.lot", [["name", "in", battery_serials]], ["id", "name"], order="name asc",
    )
    if not lots:
        return {}
    lot_ids = [lot["id"] for lot in lots]

    # Current location per lot: the location holding the highest quantity,
    # same tie-break as battery_warranty_etl.fetch_current_locations.
    quants = search_read_all(
        "stock.quant", [["lot_id", "in", lot_ids], ["quantity", ">", 0]],
        ["lot_id", "location_id", "quantity"], order="lot_id asc",
    )
    loc_ids = list({q["location_id"][0] for q in quants if q.get("location_id")})
    locs = search_read_all(
        "stock.location", [["id", "in", loc_ids]], ["id", "complete_name", "usage"],
    ) if loc_ids else []
    loc_by_id = {loc["id"]: loc for loc in locs}

    loc_map = {}
    for q in quants:
        lot_id = q["lot_id"][0]
        loc = loc_by_id.get(q["location_id"][0], {})
        existing = loc_map.get(lot_id)
        if not existing or q["quantity"] > existing["quantity"]:
            loc_map[lot_id] = {
                "location_id": q["location_id"][0],
                "location_name": loc.get("complete_name", ""),
                "usage": loc.get("usage", ""),
            }

    # For lots currently at a customer, trace back to who it was delivered
    # to (same pattern as compute_first_delivery_to_current_owner).
    customer_lot_ids = [lid for lid in lot_ids if loc_map.get(lid, {}).get("usage") == "customer"]
    deliveries_by_lot = {}
    if customer_lot_ids:
        lines = search_read_all(
            "stock.move.line",
            [
                ["lot_id", "in", customer_lot_ids],
                ["picking_id.picking_type_id.code", "=", "outgoing"],
                ["picking_id.state", "=", "done"],
                ["location_dest_id.usage", "=", "customer"],
            ],
            ["lot_id", "location_dest_id", "picking_id", "date"], order="date asc",
        )
        for ml in lines:
            deliveries_by_lot.setdefault(ml["lot_id"][0], []).append({
                "location_dest_id": ml["location_dest_id"][0],
                "picking_id": ml["picking_id"][0],
                "date": str(ml.get("date", ""))[:10],
            })

    first_delivery_map = {}
    for lot_id, records in deliveries_by_lot.items():
        loc = loc_map.get(lot_id)
        if not loc or loc["usage"] != "customer":
            continue
        matching = [r for r in records if r["location_dest_id"] == loc["location_id"]] or records
        matching.sort(key=lambda r: r["date"])
        first_delivery_map[lot_id] = matching[0]

    # Resolve to the actual customer company (commercial_partner_id), not
    # the raw delivery contact — Odoo's partner_id on a picking is often a
    # specific ship-to contact/address under that customer, whose own
    # display_name is a "<Company>, <Contact>" compound (e.g. "EIE
    # Industrial power a division of CFAO Equipment SA, Durban"). That's
    # not a customer name, it's a location description wearing a comma.
    # commercial_partner_id walks up to the actual top-level customer
    # regardless of which specific contact/address was on the delivery.
    picking_ids = list({d["picking_id"] for d in first_delivery_map.values()})
    picking_partner = {}
    picking_customer_ref = {}
    if picking_ids:
        pickings = search_read_all(
            "stock.picking", [["id", "in", picking_ids]], ["id", "partner_id", "sale_id"],
        )
        partner_ids = list({p["partner_id"][0] for p in pickings if p.get("partner_id")})
        customer_name_by_partner_id = {}
        if partner_ids:
            partners = search_read_all(
                "res.partner", [["id", "in", partner_ids]], ["id", "commercial_partner_id"],
            )
            customer_name_by_partner_id = {
                p["id"]: (p["commercial_partner_id"][1] if p.get("commercial_partner_id") else "")
                for p in partners
            }
        picking_partner = {
            p["id"]: customer_name_by_partner_id.get(p["partner_id"][0], "") if p.get("partner_id") else ""
            for p in pickings
        }

        # Customer Reference (client_order_ref) lives on the linked Sale
        # Order, not on the picking itself — stock.picking.sale_id gives a
        # direct link, no need to parse the picking's free-text 'origin'.
        sale_ids = list({p["sale_id"][0] for p in pickings if p.get("sale_id")})
        customer_ref_by_sale_id = {}
        if sale_ids:
            orders = search_read_all(
                "sale.order", [["id", "in", sale_ids]], ["id", "client_order_ref"],
            )
            customer_ref_by_sale_id = {
                o["id"]: (_trim_po_prefix(o["client_order_ref"]) if o.get("client_order_ref") else None)
                for o in orders
            }
        picking_customer_ref = {
            p["id"]: customer_ref_by_sale_id.get(p["sale_id"][0]) if p.get("sale_id") else None
            for p in pickings
        }

    resolved = {}
    for lot in lots:
        lot_id = lot["id"]
        loc = loc_map.get(lot_id, {})
        deliv = first_delivery_map.get(lot_id) if loc.get("usage") == "customer" else None
        picking_id = deliv["picking_id"] if deliv else None

        if loc.get("usage") == "customer":
            name = (picking_partner.get(picking_id) if picking_id else None) or loc.get("location_name") or None
        else:
            name = loc.get("location_name") or None

        customer_reference = picking_customer_ref.get(picking_id) if picking_id else None

        if name:
            resolved[lot["name"]] = {"name": name, "customer_reference": customer_reference}

    return resolved


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
# Ageing / historical-overdue flag
# ---------------------------------------------------------------------------

def add_ageing_column(rows: list, has_header: bool, historical_age_days: float = None) -> list:
    """
    Appends 'Days Since Started' (integer day count) based on the
    'Started At' column, then re-sorts the data rows so historical/
    overdue alerts (see HISTORICAL_AGE_DAYS) surface at the top instead
    of sitting buried at the bottom of an otherwise reverse-chronological
    list. write_data_sheet re-derives which rows are historical straight
    from this column (value >= historical_age_days) to apply the red
    highlight fill — there's no separate text flag column, since between
    the sort position, the highlight, and the day count itself, a fourth
    "Historical — overdue" label added nothing.

    This is a ROLLING threshold measured from today, not a fixed
    calendar cutoff — so it keeps meaning the same thing ("this has been
    sitting unresolved a long time") no matter how many weeks this
    report keeps running for, instead of needing the cutoff date
    hand-updated each time.

    The re-sort is stable: within the "historical" group and the
    "current" group, rows keep their existing relative order (normally
    reverse-chronological), so this only pulls the historical block to
    the front rather than fully re-ordering everything.

    Rows with a blank/unparseable 'Started At' are left blank in the new
    column and sort as non-historical (i.e. they stay wherever they'd
    otherwise land, rather than being pulled to the top).
    """
    if not has_header:
        print("  (Skipping Ageing: --no-header means there's no 'Started At' column to key off)")
        return rows

    headers = rows[0]
    if "Started At" not in headers:
        print("  (Skipping Ageing: no 'Started At' column found)")
        return rows

    if historical_age_days is None:
        historical_age_days = HISTORICAL_AGE_DAYS

    started_col = headers.index("Started At")
    today = datetime.now().date()

    enriched = []
    for row in rows[1:]:
        row = list(row)
        started_raw = row[started_col].strip() if started_col < len(row) and row[started_col] else ""
        dt = parse_datetime(started_raw)
        if dt is None:
            days_since = ""
            is_historical = False
        else:
            days_since = (today - dt.date()).days
            is_historical = days_since >= historical_age_days
        row.append(days_since)
        enriched.append((is_historical, row))

    historical_count = sum(1 for is_historical, _ in enriched if is_historical)
    print(f"  Ageing: {historical_count}/{len(enriched)} alerts are at least "
          f"{historical_age_days:g} days old — moved to the top")

    enriched.sort(key=lambda pair: 0 if pair[0] else 1)

    new_header = list(rows[0]) + ["Days Since Started"]
    return [new_header] + [row for _, row in enriched]


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
      - There's also an upper bound, `time_stamp <= <now>`. batteries_v3
        has been observed (e.g. device P16-000097) returning corrupt rows
        stamped with impossible future dates like 2112-01-01. Without an
        upper bound, ORDER BY time_stamp DESC always ranks such a row
        first, permanently shadowing the device's real recent reading.
        Real readings can't be timestamped in the future, so this filter
        never excludes a row we'd actually want.

    Returns {bem_id: (soc_str, timestamp_str)}. bem_ids not found within
    the lookback window are simply absent from the returned dict.
    """
    if not bem_ids:
        return {}

    if lookback_hours is None:
        lookback_hours = SOC_LOOKBACK_HOURS

    now = datetime.now(timezone.utc).replace(tzinfo=None)
    cutoff = now - timedelta(hours=lookback_hours)
    now_literal = now.strftime("%Y-%m-%d %H:%M:%S")
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
              AND {ATHENA_TIMESTAMP_COL} <= timestamp '{now_literal}'
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


def get_last_seen(
    bem_ids: list,
    aws_profile: str = None,
    region: str = None,
    s3_output: str = None,
) -> dict:
    """
    For each bem_id in `bem_ids`, finds the single most recent time_stamp
    it has EVER reported — no lookback window (unlike get_recent_soc/
    get_soc_at_timestamps, deliberately no hour_inserted/time_stamp LOWER
    bound, since the whole point is to find readings that may be far
    outside any lookback window). This is what lets a battery that has
    simply gone quiet (present here, with an old timestamp) be told apart
    from one that has never reported at all (absent from the result
    entirely).

    Does still cap time_stamp at "now": batteries_v3 has been observed
    (e.g. device P16-000097) returning rows corrupted with impossible
    future dates like 2112-01-01, and MAX(time_stamp) with no upper bound
    would report that garbage value as the "last seen" reading. A real
    reading can't be timestamped in the future, so this never excludes a
    row we'd actually want.

    Only call this for the (typically small) set of bem_ids that already
    missed a windowed lookup elsewhere, so the full-history scan stays
    cheap.

    Returns {bem_id: last_seen_timestamp_str}. bem_ids with no reading in
    the table at all are simply absent from the returned dict.
    """
    if not bem_ids:
        return {}

    now_literal = datetime.now(timezone.utc).replace(tzinfo=None).strftime("%Y-%m-%d %H:%M:%S")
    quoted_ids = ", ".join("'" + bid.replace("'", "''") + "'" for bid in bem_ids)

    sql = f"""
        SELECT {ATHENA_BEM_ID_COL} AS bem_id, MAX({ATHENA_TIMESTAMP_COL}) AS last_seen
        FROM {ATHENA_DATABASE}.{ATHENA_TABLE}
        WHERE {ATHENA_BEM_ID_COL} IN ({quoted_ids})
          AND {ATHENA_TIMESTAMP_COL} <= timestamp '{now_literal}'
        GROUP BY {ATHENA_BEM_ID_COL}
    """

    rows = _run_athena_query(sql, aws_profile=aws_profile, region=region, s3_output=s3_output)

    last_seen_by_id = {}
    for r in rows:
        bem_id = r.get("bem_id")
        if bem_id:
            last_seen_by_id[bem_id] = r.get("last_seen")

    return last_seen_by_id


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
    on the 'Battery' column (treated as bem_id). Blank/placeholder ('-')
    Battery values are left blank in both columns. Batteries with a real
    id but no reading in Athena within the lookback window get a
    no-data placeholder in both columns instead — "Battery
    Offline/Disconnected" if they have never reported at all, or
    "No Communication Since <timestamp>" (their last-ever-seen reading)
    if they reported before and went quiet.
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

    missing_ids = [bid for bid in bem_ids if bid not in soc_by_id]
    last_seen_by_id = {}
    if missing_ids:
        print(f"  Recent SoC: checking last-ever-seen timestamp for "
              f"{len(missing_ids)} batteries with no reading in the last "
              f"{effective_lookback:g}h")
        last_seen_by_id = get_last_seen(
            missing_ids, aws_profile=aws_profile, region=region, s3_output=s3_output,
        )

    new_rows = []
    for r_idx, row in enumerate(rows):
        row = list(row)
        if r_idx == 0:
            row.append("Recent SoC (%)")
            row.append("Recent SoC Updated At")
        else:
            value = row[battery_col].strip() if battery_col < len(row) and row[battery_col] else ""
            soc, ts = soc_by_id.get(value, (None, None))
            if soc is not None:
                row.append(soc)
                row.append(ts)
            elif value and value != "-":
                last_seen = last_seen_by_id.get(value)
                placeholder = f"{NO_COMM_PREFIX}{last_seen}" if last_seen else BATTERY_OFFLINE_TEXT
                row.append(placeholder)
                row.append(placeholder)
            else:
                row.append("")
                row.append("")
        new_rows.append(row)

    return new_rows


def add_stuck_soc_flag(
    rows: list,
    has_header: bool,
    soc_threshold_pct: float = None,
    min_days: float = None,
) -> list:
    """
    Appends a 'SoC Health' column, flagging a row when its battery's
    'Recent SoC (%)' is at/below `soc_threshold_pct` (default
    STUCK_SOC_THRESHOLD_PCT, ~5%) AND the alert has been open at least
    `min_days` (default STUCK_SOC_MIN_DAYS, ~14 days) — i.e. a battery
    that's been sitting dead for a long time rather than one that's just
    started draining.

    This is a heuristic from the two numbers already on hand (today's
    Recent SoC, and today minus Started At), not a scan of the battery's
    full SoC history — it can't tell "flat at 0% the entire time" apart
    from "recovered partway through and dropped back to 0% again". If
    that distinction matters, it needs its own Athena query (MIN/MAX SoC
    over the alert's lifetime) rather than reusing this column.

    Must run after add_recent_soc_column (needs 'Recent SoC (%)') and
    requires 'Battery' and 'Started At'. Rows with a non-numeric Recent
    SoC (blank, or one of the add_recent_soc_column placeholders like
    "Battery Offline/Disconnected") or an unparseable Started At are left
    blank — there's nothing to compare against.
    """
    if not has_header:
        print("  (Skipping SoC Health: --no-header means there's no 'Recent SoC (%)'/'Started At' column to key off)")
        return rows

    headers = rows[0]
    required = ("Battery", "Started At", "Recent SoC (%)")
    if not all(h in headers for h in required):
        print("  (Skipping SoC Health: need 'Battery', 'Started At', and 'Recent SoC (%)' columns)")
        return rows

    if soc_threshold_pct is None:
        soc_threshold_pct = STUCK_SOC_THRESHOLD_PCT
    if min_days is None:
        min_days = STUCK_SOC_MIN_DAYS

    started_col = headers.index("Started At")
    soc_col = headers.index("Recent SoC (%)")
    today = datetime.now().date()

    new_rows = [rows[0] + ["SoC Health"]]
    flagged = 0
    for row in rows[1:]:
        row = list(row)
        started_raw = row[started_col].strip() if started_col < len(row) and row[started_col] else ""
        soc_raw = row[soc_col] if soc_col < len(row) else None

        dt = parse_datetime(started_raw)
        status = ""
        if dt is not None:
            try:
                soc_value = float(soc_raw)
            except (TypeError, ValueError):
                soc_value = None
            if soc_value is not None:
                days_open = (today - dt.date()).days
                if soc_value <= soc_threshold_pct and days_open >= min_days:
                    status = f"Stuck at {soc_value:g}% — No Recovery ({days_open}d)"
                    flagged += 1
        row.append(status)
        new_rows.append(row)

    print(f"  SoC Health: {flagged} alert(s) flagged as stuck at "
          f"<={soc_threshold_pct:g}% SoC for >={min_days:g} days with no recovery")

    return new_rows


def get_soc_at_timestamps(
    pairs: list,
    aws_profile: str = None,
    region: str = None,
    s3_output: str = None,
    search_window_hours: float = None,
) -> dict:
    """
    For each (device_serial, target_dt) pair in `pairs`, finds the most
    recent SoC reading at or before target_dt, searching back up to
    `search_window_hours` (default SOC_AT_START_WINDOW_HOURS) in case
    there's a gap right at that moment. Uses a single Athena query built
    from a VALUES list covering every pair at once.

    Each pair is matched back by a synthetic integer index rather than by
    round-tripping the timestamp through Athena's string formatting (which
    can add trailing zeros/microseconds), so there's no ambiguity matching
    results back to the original (device_serial, target_dt) keys.

    Returns {(device_serial, target_dt): (soc_str, time_stamp_str)}. Pairs
    with no reading found in the window are simply absent from the dict.
    """
    if not pairs:
        return {}

    if search_window_hours is None:
        search_window_hours = SOC_AT_START_WINDOW_HOURS
    window_minutes = max(1, round(search_window_hours * 60))

    values_list = ",\n            ".join(
        f"({idx}, '{dev.replace(chr(39), chr(39) * 2)}', "
        f"TIMESTAMP '{dt.strftime('%Y-%m-%d %H:%M:%S')}')"
        for idx, (dev, dt) in enumerate(pairs)
    )

    # Global bounds across all targets, used only to help Athena/Iceberg
    # skip files — correctness comes from the per-row
    # b.time_stamp <= t.target_ts join condition below, not from these.
    earliest_target = min(dt for _, dt in pairs)
    latest_target = max(dt for _, dt in pairs)
    global_lower = earliest_target - timedelta(hours=search_window_hours)
    global_hour_inserted_floor = global_lower.strftime("%Y-%m-%d-%H")
    global_time_stamp_floor = global_lower.strftime("%Y-%m-%d %H:%M:%S")
    global_time_stamp_ceiling = latest_target.strftime("%Y-%m-%d %H:%M:%S")

    sql = f"""
        WITH targets (idx, device_serial, target_ts) AS (
            VALUES
            {values_list}
        ),
        candidates AS (
            SELECT
                t.idx,
                b.bem_id,
                b.sequence_number,
                max_by(b.{ATHENA_SOC_COL}, b.time_received) AS soc,
                CAST(max_by(b.{ATHENA_TIMESTAMP_COL}, b.time_received) AS timestamp) AS time_stamp
            FROM targets t
            JOIN {ATHENA_DATABASE}.{ATHENA_TABLE} b
              ON b.{ATHENA_BEM_ID_COL} = t.device_serial
             AND b.{ATHENA_TIMESTAMP_COL} <= t.target_ts
             AND b.{ATHENA_TIMESTAMP_COL} >= t.target_ts - INTERVAL '{window_minutes}' MINUTE
             AND b.{ATHENA_HOUR_INSERTED_COL} >= '{global_hour_inserted_floor}'
             AND b.{ATHENA_TIMESTAMP_COL} >= timestamp '{global_time_stamp_floor}'
             AND b.{ATHENA_TIMESTAMP_COL} <= timestamp '{global_time_stamp_ceiling}'
            GROUP BY t.idx, b.bem_id, b.sequence_number
        ),
        ranked AS (
            SELECT *,
                ROW_NUMBER() OVER (
                    PARTITION BY idx
                    ORDER BY time_stamp DESC, sequence_number DESC
                ) AS rn
            FROM candidates
        )
        SELECT idx, soc, time_stamp
        FROM ranked
        WHERE rn = 1
    """

    rows = _run_athena_query(sql, aws_profile=aws_profile, region=region, s3_output=s3_output)

    result = {}
    for r in rows:
        idx_raw = r.get("idx")
        if idx_raw is None:
            continue
        idx = int(idx_raw)
        if 0 <= idx < len(pairs):
            result[pairs[idx]] = (r.get("soc"), r.get("time_stamp"))

    return result


def add_soc_at_alert_time_column(
    rows: list,
    has_header: bool,
    aws_profile: str = None,
    region: str = None,
    s3_output: str = None,
    search_window_hours: float = None,
) -> list:
    """
    Inserts a 'SoC at Start (%)' column immediately after 'Started At':
    for each alert row, looks up the SoC reading in effect at (or just
    before) THAT row's own Started At timestamp — not the battery's
    latest reading (see add_recent_soc_column for that), but the reading
    closest to when that specific alert began.

    Rows with a blank/'-' Battery, a blank Started At, or an unparseable
    Started At are left blank. Rows with a real Battery/Started At but no
    reading found in the search window get a no-data placeholder instead
    — "Battery Offline/Disconnected" if that battery has never reported
    at all, or "No Communication Since <timestamp>" (its last-ever-seen
    reading) if it reported before and went quiet. Requires both
    'Battery' and 'Started At' columns to be present.
    """
    if not has_header:
        print("  (Skipping SoC at Start: --no-header means there's no 'Battery'/'Started At' column to key off)")
        return rows

    headers = rows[0]
    if BATTERY_COLUMN_NAME not in headers or "Started At" not in headers:
        print("  (Skipping SoC at Start: need both 'Battery' and 'Started At' columns)")
        return rows

    battery_col = headers.index(BATTERY_COLUMN_NAME)
    started_col = headers.index("Started At")
    data_rows = rows[1:]

    # Build the (device_serial, target_dt) each row needs, and the
    # deduplicated set of unique pairs to actually query Athena for.
    row_targets = []
    seen_pairs = {}
    for row in data_rows:
        battery = row[battery_col].strip() if battery_col < len(row) and row[battery_col] else ""
        started_raw = row[started_col].strip() if started_col < len(row) and row[started_col] else ""
        if not battery or battery == "-" or not started_raw:
            row_targets.append(None)
            continue
        target_dt = parse_datetime(started_raw)
        if target_dt is None:
            row_targets.append(None)
            continue
        pair = (battery, target_dt)
        row_targets.append(pair)
        seen_pairs[pair] = True

    unique_pairs = list(seen_pairs.keys())

    if not unique_pairs:
        print("  (Skipping SoC at Start: no rows had both a usable Battery and a parseable Started At)")
        soc_by_pair = {}
    else:
        effective_window = SOC_AT_START_WINDOW_HOURS if search_window_hours is None else search_window_hours
        print(f"  SoC at Start: looking up {len(unique_pairs)} unique battery/time "
              f"combinations from Athena (searching up to {effective_window:g}h "
              f"before each alert) — note: analytics.batteries_v3 has a known "
              f"compaction/small-files issue that can make even trivial "
              f"queries slow; this may take a while")
        soc_by_pair = get_soc_at_timestamps(
            unique_pairs, aws_profile=aws_profile, region=region, s3_output=s3_output,
            search_window_hours=effective_window,
        )
        matched = sum(1 for p in unique_pairs if p in soc_by_pair)
        print(f"  SoC at Start: matched {matched}/{len(unique_pairs)} lookups")

    missing_batteries = sorted({
        pair[0] for pair in unique_pairs
        if soc_by_pair.get(pair, (None, None))[0] is None
    })
    last_seen_by_id = {}
    if missing_batteries:
        print(f"  SoC at Start: checking last-ever-seen timestamp for "
              f"{len(missing_batteries)} batteries with no reading found in window")
        last_seen_by_id = get_last_seen(
            missing_batteries, aws_profile=aws_profile, region=region, s3_output=s3_output,
        )

    new_rows = []
    for r_idx, row in enumerate(rows):
        row = list(row)
        if r_idx == 0:
            row.insert(started_col + 1, "SoC at Start (%)")
        else:
            target = row_targets[r_idx - 1]
            if target is None:
                soc = ""
            else:
                found = soc_by_pair.get(target)
                if found is not None and found[0] is not None:
                    soc = found[0]
                else:
                    last_seen = last_seen_by_id.get(target[0])
                    soc = f"{NO_COMM_PREFIX}{last_seen}" if last_seen else BATTERY_OFFLINE_TEXT
            row.insert(started_col + 1, soc)
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


def write_data_sheet(wb, rows: list, has_header: bool = True, historical_age_days: float = None):
    ws = wb.active
    ws.title = "Data"
    headers = rows[0] if has_header else None

    NO_DATA_COLUMNS = {"Recent SoC (%)", "Recent SoC Updated At", "SoC at Start (%)"}
    days_since_idx = headers.index("Days Since Started") if headers and "Days Since Started" in headers else None

    def is_no_data_placeholder(v):
        return v == BATTERY_OFFLINE_TEXT or (isinstance(v, str) and v.startswith(NO_COMM_PREFIX))

    def resolve_fill(r_idx, is_historical_row):
        if is_historical_row:
            return HISTORICAL_FILL
        return ALT_FILL if r_idx % 2 == 0 else None

    for r_idx, row in enumerate(rows, start=1):
        is_header_row = has_header and r_idx == 1
        # Historical/overdue rows get a highlight fill across every cell,
        # on top of already being sorted to the top by add_ageing_column —
        # belt-and-braces so they're impossible to miss either way. There's
        # no separate text flag column for this — re-derive it straight
        # from 'Days Since Started' against the same threshold add_ageing_
        # column sorted by.
        is_historical_row = False
        if not is_header_row and historical_age_days is not None and days_since_idx is not None and days_since_idx < len(row):
            days_since = row[days_since_idx]
            is_historical_row = isinstance(days_since, (int, float)) and days_since >= historical_age_days

        for c_idx, value in enumerate(row, start=1):
            header_name = headers[c_idx - 1] if headers and c_idx - 1 < len(headers) else None
            is_blank = value is None or (isinstance(value, str) and value.strip() == "")
            fill = resolve_fill(r_idx, is_historical_row)

            if is_header_row:
                cell = ws.cell(row=r_idx, column=c_idx, value=str(value).strip())
                cell.font = HEADER_FONT
                cell.fill = HEADER_FILL
                cell.alignment = HEADER_ALIGN
                continue

            # The hidden 'Date' helper column must stay a real date value
            # (or genuinely blank) for the Pivot sheet's formulas to work —
            # it's never shown to the user, so it's exempt from the
            # no-data placeholders regardless of NO_DATA_COLUMNS.
            if header_name == "Date":
                cell = ws.cell(row=r_idx, column=c_idx, value=value if value != "" else None)
                if value != "":
                    cell.number_format = "dd mmm yyyy"
                cell.font = DATA_FONT
                if fill:
                    cell.fill = fill
                continue

            # The SoC/timestamp columns get a no-data placeholder instead of
            # a raw value: "Battery Offline/Disconnected" (no history at all)
            # or "No Communication Since <ts>" (reported before, then went
            # quiet) when add_recent_soc_column/add_soc_at_alert_time_column
            # already resolved one of those; "No data provided" as a last
            # resort for rows that were genuinely blank (e.g. no valid
            # Battery id to check in the first place). Everywhere else, a
            # blank cell stays genuinely blank, same as before.
            if header_name in NO_DATA_COLUMNS and (is_blank or is_no_data_placeholder(value)):
                text = value if is_no_data_placeholder(value) else NO_DATA_TEXT
                cell = ws.cell(row=r_idx, column=c_idx, value=text)
                cell.font = NO_DATA_FONT
                if fill:
                    cell.fill = fill
                continue

            if header_name in ("Hour", "Days Since Started") and not is_blank:
                cell = ws.cell(row=r_idx, column=c_idx, value=int(value))
                cell.number_format = "0"
            elif header_name in ("Recent SoC (%)", "SoC at Start (%)") and not is_blank:
                try:
                    numeric_soc = float(value)
                    cell = ws.cell(row=r_idx, column=c_idx, value=numeric_soc)
                    cell.number_format = "0.0"
                except (TypeError, ValueError):
                    cell = ws.cell(row=r_idx, column=c_idx, value=str(value).strip())
            else:
                cell = ws.cell(
                    row=r_idx, column=c_idx,
                    value=str(value).strip() if not is_blank else None
                )

            if header_name == "SoC Health" and isinstance(value, str) and value.startswith("Stuck at "):
                cell.font = STUCK_SOC_FONT
            else:
                cell.font = DATA_FONT
            if fill:
                cell.fill = fill

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
    parser.add_argument("--no-odoo-region", action="store_true",
                        help="Skip the Odoo fallback lookup for batteries with no "
                             "Organisation Unit from the alert export (requires "
                             "ODOO_URL/ODOO_DB/ODOO_USERNAME/ODOO_API_KEY to do anything anyway)")
    parser.add_argument("--lookup", default=None,
                        help="Path to a reference workbook (e.g. IP_Scrapped_data.xlsx) "
                             "whose first sheet maps 'Organisation Unit' -> "
                             "'Parent Organisation Unit'. Used to resolve Region "
                             "before falling back to dash-splitting.")
    parser.add_argument("--no-pivot", action="store_true",
                        help="Skip automatic Pivot sheet")
    parser.add_argument("--no-soc", action="store_true",
                        help="Skip the automatic Recent SoC lookup from Athena")
    parser.add_argument("--no-soc-at-start", action="store_true",
                        help="Skip looking up the SoC reading at each alert's own Started At time")
    parser.add_argument("--soc-at-start-window-hours", type=float, default=SOC_AT_START_WINDOW_HOURS,
                        help="How many hours before each alert's Started At to search for a "
                             "reading [env: ATHENA_SOC_AT_START_WINDOW_HOURS] (default: 2)")
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
    parser.add_argument("--no-ageing", action="store_true",
                        help="Skip the 'Days Since Started' column and the "
                             "historical-overdue sort/highlight")
    parser.add_argument("--historical-age-days", type=float, default=HISTORICAL_AGE_DAYS,
                        help="Alerts at least this many days old are flagged "
                             "'Historical — overdue' and sorted to the top "
                             "[env: HISTORICAL_AGE_DAYS] (default: 75)")
    parser.add_argument("--no-stuck-soc", action="store_true",
                        help="Skip the 'SoC Health' stuck-battery flag")
    parser.add_argument("--stuck-soc-threshold-pct", type=float, default=STUCK_SOC_THRESHOLD_PCT,
                        help="Recent SoC at/below this percent counts as 'low' "
                             "[env: STUCK_SOC_THRESHOLD_PCT] (default: 5)")
    parser.add_argument("--stuck-soc-min-days", type=float, default=STUCK_SOC_MIN_DAYS,
                        help="Alert must be open at least this many days, still "
                             "at/below the SoC threshold, to be flagged as stuck "
                             "[env: STUCK_SOC_MIN_DAYS] (default: 14)")
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

    if not args.no_soc_at_start:
        rows = add_soc_at_alert_time_column(
            rows, has_header,
            aws_profile=args.aws_profile,
            region=args.region,
            s3_output=args.s3_output,
            search_window_hours=args.soc_at_start_window_hours,
        )

    region_lookup = None
    if args.lookup:
        if not os.path.isfile(args.lookup):
            sys.exit(f"Lookup file not found: {args.lookup}")
        region_lookup = load_region_lookup(args.lookup)

    odoo_fallback = {}
    odoo_configured = all([ODOO_URL, ODOO_DB, ODOO_USERNAME, ODOO_API_KEY])
    if not args.no_region and not args.no_odoo_region and has_header and odoo_configured:
        headers = rows[0]
        if "Organisation Unit" in headers and BATTERY_COLUMN_NAME in headers:
            org_col = headers.index("Organisation Unit")
            battery_col = headers.index(BATTERY_COLUMN_NAME)
            needs_odoo = sorted({
                row[battery_col].strip()
                for row in rows[1:]
                if battery_col < len(row) and row[battery_col] and row[battery_col].strip()
                and org_col < len(row) and _is_blank_or_dash(row[org_col])
            })
            if needs_odoo:
                print(f"  Region: {len(needs_odoo)} batteries have no Organisation Unit from "
                      f"the alert export — checking Odoo for a fallback")
                odoo_fallback = get_odoo_region_fallback(needs_odoo)
                matched = sum(1 for b in needs_odoo if b in odoo_fallback)
                print(f"  Region: Odoo fallback resolved {matched}/{len(needs_odoo)}")

    if not args.no_region:
        rows = add_region_column(rows, has_header, region_lookup=region_lookup, odoo_fallback=odoo_fallback)

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

    if not args.no_stuck_soc:
        rows = add_stuck_soc_flag(
            rows, has_header,
            soc_threshold_pct=args.stuck_soc_threshold_pct,
            min_days=args.stuck_soc_min_days,
        )

    if not args.no_ageing:
        rows = add_ageing_column(
            rows, has_header,
            historical_age_days=args.historical_age_days,
        )

    wb = openpyxl.Workbook()
    write_data_sheet(
        wb, rows, has_header=has_header,
        historical_age_days=None if args.no_ageing else args.historical_age_days,
    )

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