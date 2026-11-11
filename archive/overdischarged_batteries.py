#!/usr/bin/env python3
"""Find batteries observed at or below a SoC threshold for a sustained period.

Usage:
    python3 overdischarged_batteries.py
    python3 overdischarged_batteries.py --threshold 10 --min-days 7
    python3 overdischarged_batteries.py --athena --aws-profile infra_dev
    python3 overdischarged_batteries.py --athena --days 28
    python3 overdischarged_batteries.py --folder /path/to/reports --output results.csv

The date range filters report dates in workbook filenames. A battery must have
a numeric SoC at or below the threshold in every available report snapshot in its
streak, and still be at or below the threshold in the latest report. The reported
duration is the calendar-day span between the first and latest low snapshots.

With --athena, all serials in the Excel reports are queried over their complete
available telemetry history by default. Use --days to limit the lookback.
SoC episodes are cross-referenced through the end-user and organisation
hierarchy in IP_Scrapped_data.xlsx, with Odoo used to fill or refresh end-user
details for Athena matches. A telemetry gap longer than 24 hours breaks an
episode by default. The full report includes all mapped regions, with a
separate KwaZulu-Natal-only report.
"""

import argparse
import csv
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import openpyxl


DEFAULT_START_DATE = date(2026, 9, 13)
DEFAULT_THRESHOLD = 10.0
DEFAULT_MIN_DAYS = 7
REGION = "Industrial Power Kwazulu Natal"
REPORT_DATE_PATTERN = re.compile(
    r"_(\d{1,2})_([A-Za-z]+)_(\d{4})_Industrial_Power_Kwazulu_Natal\.xlsx$",
    re.IGNORECASE,
)
ANY_REPORT_DATE_PATTERN = re.compile(r"_(\d{1,2})_([A-Za-z]+)_(\d{4})")
REPORT_REGION_PATTERN = re.compile(
    r"_(Industrial_Power_[A-Za-z_]+)\.xlsx$",
    re.IGNORECASE,
)
REQUIRED_COLUMNS = {"Battery", "Region", "Recent SoC (%)"}
SUMMARY_COLUMNS = (
    "Threshold (%)",
    "Below Threshold Since",
    "Last Below Threshold Report Date",
    "Duration Below Threshold (Days)",
    "Below Threshold Report Snapshots",
    "Minimum SoC During Streak (%)",
    "Latest Source Workbook",
)
ATHENA_OUTPUT_COLUMNS = (
    "Battery",
    "Region",
    "Organisation Unit",
    "Asset",
    "Region Mapping Source",
    "Latest Excel Workbook",
    "Odoo Current Customer/Location",
    "Odoo Customer Reference",
    "Odoo Derived Region",
    "Threshold (%)",
    "Episode Start",
    "Last Low Reading",
    "Low Duration (Days)",
    "Minimum SoC (%)",
    "Low Readings",
    "Latest SoC (%)",
    "Latest Reading At",
    "Status",
)


@dataclass
class Reading:
    battery: str
    soc: float
    report_date: date
    values: dict[str, object]
    source: str


@dataclass
class BatteryRegion:
    battery: str
    region: str
    organisation_unit: str
    asset: str
    source: str
    report_date: date
    resolution_source: str


def parse_report_date(path: Path) -> date:
    match = REPORT_DATE_PATTERN.search(path.name)
    if not match:
        raise ValueError(f"Cannot read report date from workbook filename: {path}")
    try:
        return datetime.strptime("_".join(match.groups()), "%d_%B_%Y").date()
    except ValueError as exc:
        raise ValueError(f"Invalid report date in workbook filename: {path}") from exc


def normalized(value: object) -> str:
    return " ".join(str(value or "").split()).casefold()


def parse_soc(value: object) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        if text.endswith("%"):
            text = text[:-1].strip()
        try:
            return float(text)
        except ValueError:
            return None
    return None


def parse_any_report_date(path: Path) -> date | None:
    for candidate in (path.name, *(parent.name for parent in path.parents)):
        match = ANY_REPORT_DATE_PATTERN.search(candidate)
        if match:
            try:
                return datetime.strptime("_".join(match.groups()), "%d_%B_%Y").date()
            except ValueError:
                continue
    return None


def canonical_region(value: object) -> str:
    normalized_value = re.sub(r"[^a-z0-9]+", "", normalized(value))
    regions = {
        "industrialpowerkwazulunatal": "Industrial Power Kwazulu Natal",
        "industrialpowergauteng": "Industrial Power Gauteng",
        "industrialpowereasterncape": "Industrial Power Eastern Cape",
        "industrialpowerwesterncape": "Industrial Power Western Cape",
        "industrialpowerfreestate": "Industrial Power Free State",
        "industrialpowermpumalanga": "Industrial Power Mpumalanga",
        "industrialpowerlimpopo": "Industrial Power Limpopo",
        "industrialpowernorthwest": "Industrial Power North West",
        "industrialpowernortherncape": "Industrial Power Northern Cape",
        "industrialpowerdealer": "Industrial Power Dealer",
        "industrialpowernational": "Industrial Power National",
    }
    return regions.get(normalized_value, "")


def resolve_mapped_region(
    value: object,
    organisation_parents: dict[str, set[str]],
    visited: frozenset[str] = frozenset(),
) -> str:
    direct = canonical_region(value)
    if direct:
        return direct
    key = normalized(value)
    if not key or key in visited:
        return ""

    found = {
        region
        for parent in organisation_parents.get(key, set())
        if (
            region := resolve_mapped_region(
                parent, organisation_parents, visited | {key}
            )
        )
    }
    return found.pop() if len(found) == 1 else ""


def load_organisation_region_lookup(path: Path) -> dict[str, set[str]]:
    if not path.is_file():
        raise FileNotFoundError(f"Region reference workbook not found: {path}")
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    parents: dict[str, set[str]] = {}
    try:
        for sheet in workbook.worksheets:
            rows = sheet.iter_rows(values_only=True)
            header = next(rows, ())
            columns = {
                str(value).strip(): index
                for index, value in enumerate(header)
                if value is not None
            }
            if not {"Organisation Unit", "Parent Organisation Unit"} <= columns.keys():
                continue
            org_index = columns["Organisation Unit"]
            parent_index = columns["Parent Organisation Unit"]
            for row in rows:
                if org_index >= len(row) or parent_index >= len(row):
                    continue
                org, parent = row[org_index], row[parent_index]
                if org and parent:
                    parents.setdefault(normalized(org), set()).add(str(parent).strip())
    finally:
        workbook.close()
    if not parents:
        raise ValueError(
            f"No Organisation Unit/Parent Organisation Unit mappings found in {path}"
        )
    return parents


def collect_battery_regions(
    folder: Path,
    region_lookup_path: Path,
) -> tuple[dict[str, BatteryRegion], int, dict[str, set[str]]]:
    organisation_parents = load_organisation_region_lookup(region_lookup_path)
    latest: dict[str, BatteryRegion] = {}
    workbook_count = 0

    for path in sorted(folder.rglob("*.xlsx")):
        if path.name.startswith("~$") or path.resolve() == region_lookup_path.resolve():
            continue
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = workbook["Data"] if "Data" in workbook.sheetnames else workbook.worksheets[0]
            rows = sheet.iter_rows(values_only=True)
            raw_headers = next(rows, ())
            headers = [
                str(value).strip() if value is not None else ""
                for value in raw_headers
            ]
            if "Battery" not in headers:
                continue
            workbook_count += 1
            columns = {name: headers.index(name) for name in headers if name}
            report_date = parse_any_report_date(path) or date.min
            filename_region_match = REPORT_REGION_PATTERN.search(path.name)
            filename_region = (
                canonical_region(filename_region_match.group(1).replace("_", " "))
                if filename_region_match
                else ""
            )

            for row in rows:
                battery_value = row[columns["Battery"]] if columns["Battery"] < len(row) else None
                if not battery_value or str(battery_value).strip() in {"", "-"}:
                    continue
                battery = str(battery_value).strip()
                battery_key = battery.casefold()
                org = (
                    row[columns["Organisation Unit"]]
                    if "Organisation Unit" in columns
                    and columns["Organisation Unit"] < len(row)
                    else ""
                )
                row_region = (
                    row[columns["Region"]]
                    if "Region" in columns and columns["Region"] < len(row)
                    else ""
                )

                resolved_region = canonical_region(row_region)
                resolution_source = "Region column"
                if not resolved_region and filename_region:
                    resolved_region = filename_region
                    resolution_source = "Workbook filename"
                if not resolved_region and org:
                    resolved_region = resolve_mapped_region(org, organisation_parents)
                    resolution_source = "IP_Scrapped_data.xlsx"
                if not resolved_region and row_region:
                    resolved_region = resolve_mapped_region(row_region, organisation_parents)
                    resolution_source = "IP_Scrapped_data.xlsx"

                candidate = BatteryRegion(
                    battery=battery,
                    region=resolved_region,
                    organisation_unit=str(org or "").strip(),
                    asset=str(
                        row[columns["Asset"]]
                        if "Asset" in columns and columns["Asset"] < len(row)
                        else ""
                    ).strip(),
                    source=str(path.relative_to(folder)),
                    report_date=report_date,
                    resolution_source=resolution_source if resolved_region else "",
                )
                current = latest.get(battery_key)
                if current is None or candidate.report_date > current.report_date:
                    latest[battery_key] = candidate
                elif candidate.report_date == current.report_date:
                    if candidate.region and current.region and candidate.region != current.region:
                        current.region = "AMBIGUOUS"
                        current.resolution_source = "Conflicting workbook regions"
                    elif candidate.region and not current.region:
                        latest[battery_key] = candidate

        finally:
            workbook.close()

    return latest, workbook_count, organisation_parents


def build_athena_query(
    serials: list[str],
    cutoff: datetime | None,
    now: datetime,
    threshold: float,
    min_days: int,
    max_gap_hours: float,
) -> str:
    import BatteryAlerts as battery_alerts

    quoted_serials = ", ".join(
        "'" + serial.replace("'", "''") + "'" for serial in sorted(serials)
    )
    now_literal = now.strftime("%Y-%m-%d %H:%M:%S")
    latest_cutoff_literal = (
        now - timedelta(hours=max_gap_hours)
    ).strftime("%Y-%m-%d %H:%M:%S")
    min_seconds = min_days * 24 * 60 * 60
    gap_seconds = int(max_gap_hours * 60 * 60)
    history_filter = ""
    episode_duration_start = "MIN(time_stamp)"
    episode_duration_filter = ""
    if cutoff is not None:
        history_start = cutoff - timedelta(days=min_days)
        cutoff_literal = cutoff.strftime("%Y-%m-%d %H:%M:%S")
        history_start_literal = history_start.strftime("%Y-%m-%d %H:%M:%S")
        history_hour_literal = history_start.strftime("%Y-%m-%d-%H")
        history_filter = f"""
              AND {battery_alerts.ATHENA_HOUR_INSERTED_COL} >= '{history_hour_literal}'
              AND {battery_alerts.ATHENA_TIMESTAMP_COL} >= timestamp '{history_start_literal}'"""
        episode_duration_start = (
            f"GREATEST(MIN(time_stamp), timestamp '{cutoff_literal}')"
        )
        episode_duration_filter = (
            f"MAX(time_stamp) >= timestamp '{cutoff_literal}' AND "
        )

    return f"""
        WITH deduplicated AS (
            SELECT
                {battery_alerts.ATHENA_BEM_ID_COL} AS device_serial,
                sequence_number,
                TRY_CAST(max_by({battery_alerts.ATHENA_SOC_COL}, time_received) AS double) AS soc,
                CAST(max_by({battery_alerts.ATHENA_TIMESTAMP_COL}, time_received) AS timestamp) AS time_stamp
            FROM {battery_alerts.ATHENA_DATABASE}.{battery_alerts.ATHENA_TABLE}
            WHERE {battery_alerts.ATHENA_BEM_ID_COL} IN ({quoted_serials})
              {history_filter}
              AND {battery_alerts.ATHENA_TIMESTAMP_COL} <= timestamp '{now_literal}'
            GROUP BY {battery_alerts.ATHENA_BEM_ID_COL}, bem_id, sequence_number
        ),
        ordered AS (
            SELECT *,
                LAG(time_stamp) OVER (
                    PARTITION BY device_serial
                    ORDER BY time_stamp, sequence_number
                ) AS previous_time_stamp,
                LAG(soc) OVER (
                    PARTITION BY device_serial
                    ORDER BY time_stamp, sequence_number
                ) AS previous_soc
            FROM deduplicated
            WHERE soc IS NOT NULL
        ),
        marked AS (
            SELECT *,
                CASE
                    WHEN previous_time_stamp IS NULL
                      OR previous_soc > {threshold}
                      OR date_diff('second', previous_time_stamp, time_stamp) > {gap_seconds}
                    THEN 1 ELSE 0
                END AS starts_segment
            FROM ordered
        ),
        segmented AS (
            SELECT *,
                SUM(starts_segment) OVER (
                    PARTITION BY device_serial
                    ORDER BY time_stamp, sequence_number
                    ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
                ) AS segment_id
            FROM marked
        ),
        low_episodes AS (
            SELECT
                device_serial,
                segment_id,
                MIN(time_stamp) AS first_low_reading,
                MAX(time_stamp) AS last_low_reading,
                MIN(soc) AS minimum_soc,
                COUNT(*) AS low_reading_count
            FROM segmented
            WHERE soc BETWEEN 0 AND {threshold}
            GROUP BY device_serial, segment_id
            HAVING {episode_duration_filter}date_diff(
                    'second',
                    {episode_duration_start},
                    MAX(time_stamp)
                  ) >= {min_seconds}
        ),
        latest AS (
            SELECT device_serial, soc AS latest_soc, time_stamp AS latest_reading_at, segment_id,
                   ROW_NUMBER() OVER (
                       PARTITION BY device_serial
                       ORDER BY time_stamp DESC, sequence_number DESC
                   ) AS row_number
            FROM segmented
        )
        SELECT
            episode.device_serial,
            episode.first_low_reading,
            episode.last_low_reading,
            episode.minimum_soc,
            episode.low_reading_count,
            latest.latest_soc,
            latest.latest_reading_at,
            CASE
                WHEN latest.latest_reading_at < timestamp '{latest_cutoff_literal}'
                    THEN 'NO RECENT TELEMETRY'
                WHEN latest.latest_soc > {threshold}
                    THEN 'RECOVERED'
                WHEN latest.segment_id = episode.segment_id
                    THEN 'ONGOING'
                ELSE 'LOW AGAIN AFTER DATA GAP'
            END AS current_status
        FROM low_episodes AS episode
        JOIN latest
          ON latest.device_serial = episode.device_serial
         AND latest.row_number = 1
        ORDER BY episode.device_serial, episode.first_low_reading
    """


def collect_athena_episodes(
    folder: Path,
    region_lookup_path: Path,
    days: int | None,
    threshold: float,
    min_days: int,
    max_gap_hours: float,
    aws_profile: str | None,
    aws_region: str | None,
) -> tuple[list[dict[str, object]], int, int]:
    import BatteryAlerts as battery_alerts

    battery_regions, workbook_count, organisation_parents = collect_battery_regions(
        folder, region_lookup_path
    )
    if not battery_regions:
        raise ValueError(f"No battery serials found in Excel reports under {folder}")

    now = datetime.now(timezone.utc).replace(tzinfo=None)
    cutoff = now - timedelta(days=days) if days is not None else None
    query = build_athena_query(
        [item.battery for item in battery_regions.values()],
        cutoff,
        now,
        threshold,
        min_days,
        max_gap_hours,
    )
    results = battery_alerts._run_athena_query(
        query,
        aws_profile=aws_profile or battery_alerts.DEFAULT_AWS_PROFILE,
        region=aws_region or battery_alerts.DEFAULT_AWS_REGION,
        s3_output=battery_alerts.DEFAULT_S3_OUTPUT,
    )
    odoo_by_battery = {}
    if results:
        try:
            odoo_by_battery = {
                str(serial).strip().casefold(): info
                for serial, info in battery_alerts.get_odoo_region_fallback(
                    sorted({
                        str(row["device_serial"]).strip()
                        for row in results
                        if row.get("device_serial")
                    })
                ).items()
            }
        except KeyError as exc:
            if exc.args != ("quantity",):
                raise
            print(
                "Warning: Odoo customer/location enrichment was skipped because "
                "BatteryAlerts.get_odoo_region_fallback raised KeyError('quantity'). "
                "Continuing with workbook-based region mapping.",
                file=sys.stderr,
            )

    episodes: list[dict[str, object]] = []
    for row in results:
        battery = str(row.get("device_serial") or "").strip()
        region_info = battery_regions.get(battery.casefold())
        odoo_info = odoo_by_battery.get(battery.casefold(), {})
        odoo_name = str(odoo_info.get("name") or "").strip()
        odoo_reference = str(odoo_info.get("customer_reference") or "").strip()
        odoo_region = (
            resolve_mapped_region(odoo_name, organisation_parents)
            if odoo_name
            else ""
        )
        if not odoo_region and odoo_reference:
            odoo_region = battery_alerts.region_from_sold_to_customer(odoo_reference)
        workbook_region = region_info.region if region_info else "Not mapped"
        resolved_region = odoo_region or workbook_region
        mapping_source = (
            "Odoo current end-user/location"
            if odoo_region
            else region_info.resolution_source if region_info else ""
        )
        first_low_reading = datetime.fromisoformat(
            str(row["first_low_reading"])
        )
        last_low_reading = datetime.fromisoformat(str(row["last_low_reading"]))
        duration_start = (
            max(first_low_reading, cutoff) if cutoff is not None else first_low_reading
        )
        episodes.append(
            {
                "Battery": battery,
                "Region": resolved_region,
                "Organisation Unit": region_info.organisation_unit if region_info else "",
                "Asset": region_info.asset if region_info else "",
                "Region Mapping Source": mapping_source,
                "Latest Excel Workbook": region_info.source if region_info else "",
                "Odoo Current Customer/Location": odoo_name,
                "Odoo Customer Reference": odoo_reference,
                "Odoo Derived Region": odoo_region,
                "Threshold (%)": threshold,
                "Episode Start": first_low_reading,
                "Last Low Reading": last_low_reading,
                "Low Duration (Days)": max(
                    0,
                    (last_low_reading - duration_start).total_seconds()
                    / 86400,
                ),
                "Minimum SoC (%)": row.get("minimum_soc", ""),
                "Low Readings": row.get("low_reading_count", ""),
                "Latest SoC (%)": row.get("latest_soc", ""),
                "Latest Reading At": row.get("latest_reading_at", ""),
                "Status": row.get("current_status", ""),
            }
        )
    episodes.sort(
        key=lambda item: (
            str(item["Region"]).casefold(),
            str(item["Battery"]).casefold(),
            str(item["Episode Start"]),
        )
    )
    return episodes, len(battery_regions), workbook_count


def report_files(
    folder: Path, start_date: date, end_date: date
) -> list[tuple[Path, date]]:
    reports = []
    for path in folder.rglob("*Industrial_Power_Kwazulu_Natal.xlsx"):
        if path.name.startswith("~$"):
            continue
        report_date = parse_report_date(path)
        if start_date <= report_date <= end_date:
            reports.append((path, report_date))
    return sorted(reports, key=lambda item: (item[1], str(item[0]).casefold()))


def collect_streaks(
    reports: list[tuple[Path, date]],
    threshold: float,
    min_days: int,
    folder: Path,
) -> tuple[list[str], list[dict[str, object]]]:
    headers: list[str] = []
    active_streaks: dict[str, list[Reading]] = {}
    known_batteries: set[str] = set()

    for path, report_date in reports:
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            if "Data" not in workbook.sheetnames:
                raise ValueError(f"Missing 'Data' sheet in workbook: {path}")
            rows = workbook["Data"].iter_rows(values_only=True)
            try:
                raw_headers = next(rows)
            except StopIteration as exc:
                raise ValueError(f"Empty 'Data' sheet in workbook: {path}") from exc

            sheet_headers = [
                str(value).strip() if value is not None else ""
                for value in raw_headers
            ]
            missing = REQUIRED_COLUMNS - set(sheet_headers)
            if missing:
                raise ValueError(
                    f"Missing required columns in {path}: {', '.join(sorted(missing))}"
                )
            for header in sheet_headers:
                if header and header not in headers:
                    headers.append(header)
            column = {name: sheet_headers.index(name) for name in REQUIRED_COLUMNS}

            readings: dict[str, Reading] = {}
            for row in rows:
                battery_value = row[column["Battery"]]
                region = row[column["Region"]]
                if not battery_value or normalized(region) != REGION.casefold():
                    continue
                battery = str(battery_value).strip()
                battery_key = battery.casefold()
                values = {
                    header: row[index] if index < len(row) else None
                    for index, header in enumerate(sheet_headers)
                    if header
                }
                soc = parse_soc(values.get("Recent SoC (%)"))
                known_batteries.add(battery_key)
                if soc is not None:
                    readings[battery_key] = Reading(
                        battery=battery,
                        soc=soc,
                        report_date=report_date,
                        values=values,
                        source=str(path.relative_to(folder)),
                    )
        finally:
            workbook.close()

        for battery_key in known_batteries:
            reading = readings.get(battery_key)
            if reading is not None and reading.soc <= threshold:
                active_streaks.setdefault(battery_key, []).append(reading)
            else:
                active_streaks.pop(battery_key, None)

    output_rows = []
    latest_report_date = reports[-1][1]
    for streak in active_streaks.values():
        first_date = streak[0].report_date
        last_reading = streak[-1]
        duration = (last_reading.report_date - first_date).days
        if last_reading.report_date != latest_report_date or duration < min_days:
            continue
        output_rows.append(
            {
                **last_reading.values,
                "Threshold (%)": threshold,
                "Below Threshold Since": first_date.isoformat(),
                "Last Below Threshold Report Date": last_reading.report_date.isoformat(),
                "Duration Below Threshold (Days)": duration,
                "Below Threshold Report Snapshots": len(streak),
                "Minimum SoC During Streak (%)": min(reading.soc for reading in streak),
                "Latest Source Workbook": last_reading.source,
                "_sort": (first_date, last_reading.battery.casefold()),
            }
        )
    output_rows.sort(key=lambda row: row["_sort"])
    return headers, output_rows


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Find Industrial Power Kwazulu Natal batteries observed at or below a "
            "state-of-charge threshold for a minimum number of days."
        )
    )
    parser.add_argument(
        "--folder",
        type=Path,
        default=Path(__file__).resolve().parent,
        help="Folder to search recursively (default: this script's folder).",
    )
    parser.add_argument(
        "--start-date",
        type=date.fromisoformat,
        default=DEFAULT_START_DATE,
        metavar="YYYY-MM-DD",
        help="First report date to include (default: 2026-09-13).",
    )
    parser.add_argument(
        "--end-date",
        type=date.fromisoformat,
        default=date.today(),
        metavar="YYYY-MM-DD",
        help="Last report date to include (default: today).",
    )
    parser.add_argument(
        "--threshold",
        type=float,
        default=DEFAULT_THRESHOLD,
        help="SoC must be at or below this percentage (default: 10).",
    )
    parser.add_argument(
        "--min-days",
        type=int,
        default=DEFAULT_MIN_DAYS,
        help="Minimum calendar-day span between low-SoC snapshots (default: 7).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="CSV output path (default: batteries_under_<threshold>pct_for_<days>_days_<start>_to_<end>.csv).",
    )
    parser.add_argument(
        "--athena",
        action="store_true",
        help=(
            "Query Athena for SoC history for every serial in the Excel reports, "
            "then cross-reference regions with IP_Scrapped_data.xlsx."
        ),
    )
    parser.add_argument(
        "--days",
        type=int,
        default=None,
        help=(
            "Limit Athena telemetry lookback to this many days "
            "(default: full available battery history; may scan more data)."
        ),
    )
    parser.add_argument(
        "--max-gap-hours",
        type=float,
        default=24,
        help=(
            "Break a low-SoC streak when telemetry readings are farther apart "
            "than this (default: 24 hours)."
        ),
    )
    parser.add_argument(
        "--region-lookup",
        type=Path,
        help="Organisation-to-region reference workbook (default: <folder>/IP_Scrapped_data.xlsx).",
    )
    parser.add_argument(
        "--aws-profile",
        help="AWS profile for Athena (default: AWS_PROFILE/.env or the default credential chain).",
    )
    parser.add_argument(
        "--aws-region",
        help="AWS region for Athena (default: AWS_REGION/.env or the profile setting).",
    )
    args = parser.parse_args()

    folder = args.folder.expanduser().resolve()
    if not folder.is_dir():
        parser.error(f"folder does not exist or is not a directory: {folder}")
    if args.start_date > args.end_date:
        parser.error("--start-date must not be after --end-date")
    if args.threshold <= 0 or args.threshold > 100:
        parser.error("--threshold must be greater than 0 and at most 100")
    if args.min_days < 1:
        parser.error("--min-days must be at least 1")

    if args.athena:
        if args.days is not None and args.days < 1:
            parser.error("--days must be at least 1")
        if args.max_gap_hours <= 0:
            parser.error("--max-gap-hours must be greater than 0")
        region_lookup = (
            args.region_lookup.expanduser().resolve()
            if args.region_lookup
            else folder / "IP_Scrapped_data.xlsx"
        )
        try:
            episodes, serial_count, workbook_count = collect_athena_episodes(
                folder=folder,
                region_lookup_path=region_lookup,
                days=args.days,
                threshold=args.threshold,
                min_days=args.min_days,
                max_gap_hours=args.max_gap_hours,
                aws_profile=args.aws_profile,
                aws_region=args.aws_region,
            )
        except (FileNotFoundError, ValueError) as exc:
            parser.error(str(exc))

        end_date = datetime.now(timezone.utc).date()
        if args.days is None:
            default_output_name = f"athena_low_soc_episodes_lifetime_to_{end_date}.csv"
        else:
            cutoff_date = (
                datetime.now(timezone.utc) - timedelta(days=args.days)
            ).date()
            default_output_name = (
                f"athena_low_soc_episodes_{cutoff_date}_to_{end_date}.csv"
            )
        output = args.output or folder / default_output_name
        output = output.expanduser()
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open("w", newline="", encoding="utf-8-sig") as csv_file:
            writer = csv.DictWriter(
                csv_file,
                fieldnames=ATHENA_OUTPUT_COLUMNS,
                extrasaction="ignore",
            )
            writer.writeheader()
            writer.writerows(episodes)

        kzn_episodes = [
            episode for episode in episodes
            if str(episode["Region"]).casefold() == REGION.casefold()
        ]
        kzn_output = output.with_name(f"kzn_{output.name}")
        with kzn_output.open("w", newline="", encoding="utf-8-sig") as csv_file:
            writer = csv.DictWriter(
                csv_file,
                fieldnames=ATHENA_OUTPUT_COLUMNS,
                extrasaction="ignore",
            )
            writer.writeheader()
            writer.writerows(kzn_episodes)

        regions = {}
        for episode in episodes:
            region = str(episode["Region"])
            regions[region] = regions.get(region, 0) + 1
        print(
            f"Queried Athena for {serial_count} distinct battery serials from "
            f"{workbook_count} Excel report workbooks."
        )
        print(
            f"Found {len(episodes)} low-SoC episode(s) lasting at least "
            f"{args.min_days} days in "
            f"{'the full available history' if args.days is None else f'the last {args.days} days'}."
        )
        print(
            "KwaZulu-Natal episodes: "
            f"{len(kzn_episodes)}; region totals: "
            f"{', '.join(f'{name}={count}' for name, count in sorted(regions.items())) or 'none'}."
        )
        print(
            f"Telemetry gaps over {args.max_gap_hours:g} hours break a low-SoC episode."
        )
        print(f"CSV report: {output.resolve()}")
        print(f"KwaZulu-Natal CSV report: {kzn_output.resolve()}")
        return 0

    reports = report_files(folder, args.start_date, args.end_date)
    if not reports:
        print(
            f"No Industrial Power Kwazulu Natal workbooks dated "
            f"{args.start_date} through {args.end_date} were found in {folder}.",
            file=sys.stderr,
        )
        return 1

    headers, rows = collect_streaks(reports, args.threshold, args.min_days, folder)
    output = args.output or folder / (
        f"batteries_under_{args.threshold:g}pct_for_{args.min_days}_days_"
        f"{args.start_date}_to_{args.end_date}.csv"
    )
    output = output.expanduser()
    output.parent.mkdir(parents=True, exist_ok=True)
    output_headers = headers + [name for name in SUMMARY_COLUMNS if name not in headers]
    with output.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=output_headers, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)

    print(
        f"Scanned {len(reports)} KwaZulu-Natal workbook(s) dated "
        f"{args.start_date} through {args.end_date}."
    )
    print(
        f"Found {len(rows)} battery/batteries still at or below {args.threshold:g}% "
        f"for at least {args.min_days} days in the latest report."
    )
    print(f"Latest report date: {reports[-1][1]}.")
    print(f"CSV report: {output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
