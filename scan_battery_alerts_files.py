"""
Scan existing BatteryAlerts files for battery serials
------------------------------------------------------------
Quick check against files ALREADY on disk in ~/BatteryAlerts — no Gmail
API or OAuth setup needed. Covers only what's actually present locally
(currently just July 2026's weekly reports + alerts.txt), not the full
email history — worth knowing that limitation going in.

Run this FROM the ~/BatteryAlerts folder (or edit FOLDER below).
"""

import glob
import os

import openpyxl

FOLDER = os.path.expanduser("~/BatteryAlerts")

BATTERY_SERIALS = sorted(set("""
P08-000234 P16-000045 P38-000011 BE20000394 P23-000829 BE20000410 P28-000030 BE20000762 P24-000139 P20-000121
P08-000141 P15-000610 P16-000398 BE21000124 BE21000285 P16-000326 P16-000156 BE21000215 P23-000840 BE21000776
P23-000275 P15-000540 BE20000446 P28-000018 BE19090010 P20-000194 P14-000304 P15-000557 P23-000905 P14-000303
BE20000752 P15-000552 P15-000094 P23-001137 P16-000467 P15-000559 P25-000097 P20-000051 P15-000002 P29-000149
BE20000408 P20-000050 P15-000614 P20-000162 P20-000174 P20-000081 P28-000081 P20-000131 P15-000391 P28-000257
P20-000340 P20-000020 P23-000668 BE20000554 P20-000097 P15-000041 P15-000551 P16-000182 P16-000202 P28-000193
P16-000509 P15-000512 P14-000030 P23-000740 P15-000602
""".split()))


def scan_xlsx(path, serials_to_find):
    found = set()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            for cell in row:
                if cell is None:
                    continue
                cell_str = str(cell)
                for serial in serials_to_find:
                    if serial in cell_str:
                        found.add(serial)
    wb.close()
    return found


def scan_text_file(path, serials_to_find):
    found = set()
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        content = f.read()
    for serial in serials_to_find:
        if serial in content:
            found.add(serial)
    return found


def main():
    remaining = set(BATTERY_SERIALS)
    found_in_file = {}

    xlsx_files = sorted(glob.glob(os.path.join(FOLDER, "*.xlsx")))
    txt_files = sorted(glob.glob(os.path.join(FOLDER, "*.txt")))

    print(f"Scanning {len(xlsx_files)} xlsx files and {len(txt_files)} txt files in {FOLDER}\n")

    for path in xlsx_files:
        filename = os.path.basename(path)
        found = scan_xlsx(path, remaining)
        for serial in found:
            found_in_file.setdefault(serial, []).append(filename)
        print(f"  {filename}: found {len(found)} of {len(remaining)} remaining serials")

    for path in txt_files:
        filename = os.path.basename(path)
        found = scan_text_file(path, remaining)
        for serial in found:
            found_in_file.setdefault(serial, []).append(filename)
        print(f"  {filename}: found {len(found)} of {len(remaining)} remaining serials")

    print(f"\n{'='*60}")
    print(f"RESULTS — {len(BATTERY_SERIALS)} battery serials checked against LOCAL FILES ONLY")
    print("(NOTE: this only covers files currently on disk, not your full email history)")
    print(f"{'='*60}\n")

    found_serials = [s for s in BATTERY_SERIALS if s in found_in_file]
    not_found_serials = [s for s in BATTERY_SERIALS if s not in found_in_file]

    print(f"Found in at least one local file: {len(found_serials)}")
    for serial in found_serials:
        print(f"  {serial} — in: {', '.join(found_in_file[serial])}")

    print(f"\nNOT found in any local file (checked {len(xlsx_files) + len(txt_files)} files): {len(not_found_serials)}")
    for serial in not_found_serials:
        print(f"  {serial}")


if __name__ == "__main__":
    main()