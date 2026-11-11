#!/usr/bin/env python3
"""
Fill in Region and Sold To for a list of battery serials (e.g.
Unallocate.csv) by looking each one up in Odoo.

Usage:
    python unallocated_batteries.py                       # reads Unallocate.csv
    python unallocated_batteries.py my_list.csv -o out.csv

Input: a tab- or comma-separated file whose first column is the battery
serial (header row expected, e.g. "Battery  Varient  Region  Sold to").
Any other input columns are carried through; Region / Sold to are
overwritten with what Odoo says.

For each serial, the most recent done delivery to a customer location
(outgoing stock.picking) is used:
    Sold To             commercial customer on the Sale Order
    Customer Reference  Sale Order client_order_ref (the end user, e.g.
                        "PO: 23920 KARGO NATIONAL JHB")
    Delivered To        the ship-to contact on the delivery
    Region              first hit of:
                          1. a city abbreviation (JHB, DBN, PE, CPT ... —
                             same table as BatteryAlerts.py) or city name
                             in the Customer Reference
                          2. the province on the ship-to address
                          3. the ship-to postal code (SA code ranges)
                          4. a city abbreviation/name in the ship-to city
                          5. ... or in the Delivered To name
    Region Source       which of the above produced Region

Serials that were never delivered get blank delivery columns, with their
current stock location in 'Current Location'.

Requires ODOO_URL, ODOO_DB, ODOO_USERNAME, ODOO_API_KEY (env vars, e.g.
via .env — same keys as BatteryAlerts.py).
"""

import os
import re
import sys
import csv
import argparse
from datetime import datetime

from customer_batteries import Odoo
from BatteryAlerts import region_from_sold_to_customer


# Odoo res.country.state names (normalised: lower-case letters only) -> Region.
PROVINCE_REGION = {
    "gauteng":       "Industrial Power Gauteng",
    "kwazulunatal":  "Industrial Power Kwazulu Natal",
    "easterncape":   "Industrial Power Eastern Cape",
    "westerncape":   "Industrial Power Western Cape",
    "freestate":     "Industrial Power Free State",
    "mpumalanga":    "Industrial Power Mpumalanga",
    "limpopo":       "Industrial Power Limpopo",
    "northwest":     "Industrial Power North West",
    "northerncape":  "Industrial Power Northern Cape",
}

# South African postal code ranges -> Region (first match wins).
POSTCODE_REGION = [
    (1,    299,  "Industrial Power Gauteng"),        # Pretoria / Tshwane
    (300,  499,  "Industrial Power North West"),
    (500,  999,  "Industrial Power Limpopo"),
    (1000, 1399, "Industrial Power Mpumalanga"),
    (1400, 2199, "Industrial Power Gauteng"),
    (2200, 2499, "Industrial Power Mpumalanga"),
    (2500, 2899, "Industrial Power North West"),
    (2900, 4730, "Industrial Power Kwazulu Natal"),
    (4731, 6499, "Industrial Power Eastern Cape"),
    (6500, 8099, "Industrial Power Western Cape"),
    (8100, 8999, "Industrial Power Northern Cape"),
    (9300, 9999, "Industrial Power Free State"),
]

# Full city / suburb names seen in references and delivery addresses.
CITY_REGION = {
    "JOHANNESBURG": "Industrial Power Gauteng", "JOBURG": "Industrial Power Gauteng",
    "PRETORIA": "Industrial Power Gauteng", "ISANDO": "Industrial Power Gauteng",
    "KEMPTON PARK": "Industrial Power Gauteng", "JET PARK": "Industrial Power Gauteng",
    "GERMISTON": "Industrial Power Gauteng", "ELANDSFONTEIN": "Industrial Power Gauteng",
    "MIDRAND": "Industrial Power Gauteng", "CENTURION": "Industrial Power Gauteng",
    "BOKSBURG": "Industrial Power Gauteng", "SANDTON": "Industrial Power Gauteng",
    "DURBAN": "Industrial Power Kwazulu Natal", "PINETOWN": "Industrial Power Kwazulu Natal",
    "PIETERMARITZBURG": "Industrial Power Kwazulu Natal", "KWAZULU NATAL": "Industrial Power Kwazulu Natal",
    "PORT ELIZABETH": "Industrial Power Eastern Cape", "GQEBERHA": "Industrial Power Eastern Cape",
    "EAST LONDON": "Industrial Power Eastern Cape", "P.E": "Industrial Power Eastern Cape",
    "CAPE TOWN": "Industrial Power Western Cape", "ATLANTIS": "Industrial Power Western Cape",
    "EPPING": "Industrial Power Western Cape", "MONTAGUE GARDENS": "Industrial Power Western Cape",
    "BLOEMFONTEIN": "Industrial Power Free State",
}
_CITY_RE = re.compile(
    r"(?<![A-Za-z])(" + "|".join(re.escape(c) for c in sorted(CITY_REGION, key=len, reverse=True)) + r")(?![A-Za-z])",
    re.IGNORECASE,
)

LOOKUP_CHUNK = 500  # serials per stock.lot query


def read_input(path: str):
    with open(path, encoding="utf-8-sig", errors="replace") as f:
        lines = [l.rstrip("\r\n") for l in f if l.strip()]
    if not lines:
        sys.exit("The input file is empty.")
    delim = "\t" if "\t" in lines[0] else ","
    rows = list(csv.reader(lines, delimiter=delim))
    headers = [h.strip() for h in rows[0]]
    data = [r for r in rows[1:] if r and r[0].strip()]
    return headers, data


def province_region(state_name: str) -> str:
    key = re.sub(r"[^a-z]", "", (state_name or "").lower())
    return PROVINCE_REGION.get(key, "")


def postcode_region(zip_code) -> str:
    digits = re.sub(r"\D", "", str(zip_code or ""))
    if len(digits) != 4:
        return ""
    code = int(digits)
    return next((region for lo, hi, region in POSTCODE_REGION if lo <= code <= hi), "")


def text_region(text: str) -> str:
    """Region from a city abbreviation (JHB, DBN...) or full city name in free text."""
    region = region_from_sold_to_customer(text or "")
    if not region:
        match = _CITY_RE.search(text or "")
        region = CITY_REGION[match.group(1).upper()] if match else ""
    return region


def lookup_serials(odoo: Odoo, serials: list) -> dict:
    """Returns {serial: {...odoo fields...}} for every serial found in Odoo."""
    lots = []
    for i in range(0, len(serials), LOOKUP_CHUNK):
        lots += odoo.search_read_all(
            "stock.lot", [["name", "in", serials[i:i + LOOKUP_CHUNK]]], ["id", "name"],
        )
    if not lots:
        return {}
    lot_ids = [lot["id"] for lot in lots]

    # Current location — highest-quantity quant wins, as in BatteryAlerts.
    quants = odoo.search_read_all(
        "stock.quant", [["lot_id", "in", lot_ids], ["quantity", ">", 0]],
        ["lot_id", "location_id", "quantity"],
    )
    current_loc, best_qty = {}, {}
    for q in quants:
        lot_id = q["lot_id"][0]
        if q.get("location_id") and q["quantity"] > best_qty.get(lot_id, 0):
            best_qty[lot_id] = q["quantity"]
            current_loc[lot_id] = q["location_id"]

    # Latest done delivery to a customer per lot (lines are date-ascending).
    lines = odoo.search_read_all(
        "stock.move.line",
        [
            ["lot_id", "in", lot_ids],
            ["picking_id.picking_type_id.code", "=", "outgoing"],
            ["picking_id.state", "=", "done"],
            ["location_dest_id.usage", "=", "customer"],
        ],
        ["lot_id", "picking_id", "date"], order="date asc",
    )
    last_delivery = {ml["lot_id"][0]: ml for ml in lines}

    picking_ids = list({ml["picking_id"][0] for ml in last_delivery.values()})
    pickings = odoo.search_read_all(
        "stock.picking", [["id", "in", picking_ids]], ["id", "name", "partner_id", "sale_id"],
    ) if picking_ids else []
    picking_by_id = {p["id"]: p for p in pickings}

    partner_ids = list({p["partner_id"][0] for p in pickings if p.get("partner_id")})
    partners = odoo.search_read_all(
        "res.partner", [["id", "in", partner_ids]], ["id", "display_name", "state_id", "zip", "city"],
    ) if partner_ids else []
    partner_by_id = {p["id"]: p for p in partners}

    sale_ids = list({p["sale_id"][0] for p in pickings if p.get("sale_id")})
    orders = odoo.search_read_all(
        "sale.order", [["id", "in", sale_ids]], ["id", "name", "client_order_ref", "partner_id"],
    ) if sale_ids else []
    order_by_id = {o["id"]: o for o in orders}

    # Sold To = the commercial (company) partner of the SO customer.
    so_partner_ids = list({o["partner_id"][0] for o in orders if o.get("partner_id")})
    so_partners = odoo.search_read_all(
        "res.partner", [["id", "in", so_partner_ids]], ["id", "commercial_partner_id"],
    ) if so_partner_ids else []
    company_by_partner = {
        p["id"]: p["commercial_partner_id"][1] if p.get("commercial_partner_id") else ""
        for p in so_partners
    }

    result = {}
    for lot in lots:
        lot_id = lot["id"]
        info = {
            "Sold To": "", "Customer Reference": "", "Sale Order": "",
            "Delivery": "", "Delivery Date": "", "Delivered To": "",
            "Region": "", "Region Source": "",
            "Current Location": current_loc[lot_id][1] if lot_id in current_loc else "",
        }
        ml = last_delivery.get(lot_id)
        if ml:
            picking = picking_by_id.get(ml["picking_id"][0], {})
            partner = partner_by_id.get(picking["partner_id"][0], {}) if picking.get("partner_id") else {}
            order = order_by_id.get(picking["sale_id"][0], {}) if picking.get("sale_id") else {}

            info["Delivery"] = picking.get("name", "")
            info["Delivery Date"] = str(ml.get("date", ""))[:10]
            info["Delivered To"] = partner.get("display_name", "")
            info["Sale Order"] = order.get("name", "")
            info["Customer Reference"] = order.get("client_order_ref") or ""
            if order.get("partner_id"):
                info["Sold To"] = company_by_partner.get(order["partner_id"][0]) or order["partner_id"][1]

            region, source = "", ""
            for label, value in [
                ("Customer Reference", text_region(info["Customer Reference"])),
                ("Delivery Address Province", province_region(partner["state_id"][1]) if partner.get("state_id") else ""),
                ("Delivery Address Postal Code", postcode_region(partner.get("zip"))),
                ("Delivery Address City", text_region(partner.get("city") or "")),
                ("Delivered To Name", text_region(info["Delivered To"])),
            ]:
                if value:
                    region, source = value, label
                    break
            info["Region"], info["Region Source"] = region, source

        result[lot["name"]] = info
    return result


def main():
    parser = argparse.ArgumentParser(description="Add Region / Sold To / Customer Reference from Odoo to a battery serial list.")
    parser.add_argument("input", nargs="?", default="Unallocate.csv", help="Serial list (default: Unallocate.csv)")
    parser.add_argument("-o", "--output", help="Output CSV path (default: <input>_Odoo_<date>.csv)")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        sys.exit(f"File not found: {args.input}")

    headers, data = read_input(args.input)
    serials = sorted({r[0].strip() for r in data})
    print(f"Looking up {len(serials)} serials in Odoo...")

    odoo = Odoo()
    found = lookup_serials(odoo, serials)

    # Input columns first (Region / Sold to replaced by Odoo values), then the extra Odoo detail.
    drop = {"region", "sold to"}
    keep_idx = [i for i, h in enumerate(headers) if h.lower() not in drop]
    extra = ["Region", "Sold To", "Customer Reference", "Sale Order", "Delivery",
             "Delivery Date", "Delivered To", "Region Source", "Current Location", "Found In Odoo"]
    out_headers = [headers[i] for i in keep_idx] + extra

    out_rows = []
    for r in data:
        info = found.get(r[0].strip())
        row = [r[i].strip() if i < len(r) else "" for i in keep_idx]
        row += [info[c] for c in extra[:-1]] if info else [""] * (len(extra) - 1)
        row.append("Yes" if info else "No")
        out_rows.append(row)

    base = os.path.splitext(os.path.basename(args.input))[0]
    output = args.output or f"{base}_Odoo_{datetime.now():%Y-%m-%d}.csv"
    with open(output, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(out_headers)
        writer.writerows(out_rows)

    delivered = sum(1 for i in found.values() if i["Delivery"])
    with_region = sum(1 for i in found.values() if i["Region"])
    print(f"\n{len(found)}/{len(serials)} serials found in Odoo, "
          f"{delivered} delivered to a customer, {with_region} with a Region.")
    print(f"Wrote {len(out_rows)} rows to {output}")


if __name__ == "__main__":
    main()
