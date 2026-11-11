#!/usr/bin/env python3
"""
Export every battery serial delivered to an end user/customer, with its
Sale Order, straight from Odoo.

Usage:
    python customer_batteries.py                   # prompts for the name
    python customer_batteries.py "Kargo"           # or pass it as an argument
    python customer_batteries.py "kargo jhb" -o out.csv

The end user is found via the Sale Order's Customer Reference field
(sale.order.client_order_ref), which is where the end user's name is
written — e.g. "PO: 23920 KARGO NATIONAL JHB". Matching is
case-insensitive and word-by-word: every word typed must appear somewhere
in the reference, in any order, so "kargo jhb" matches both
"PO: 17278-17280 KARGO NATIONAL - JHB" and "PO: 15396 JHB - KARGO NATIONAL".
Cancelled Sale Orders are ignored.

For each matching Sale Order, the serials come from its done warehouse-out
deliveries (stock.picking, outgoing).

Output CSV columns:
    Serial, Product, Sale Order, Customer Reference, Invoiced Customer,
    Delivery, Delivery Date, Still At Customer

'Still At Customer' is Yes when the serial's current stock location is a
customer location (i.e. it hasn't since been returned to one of our
warehouses), No otherwise.

Requires ODOO_URL, ODOO_DB, ODOO_USERNAME, ODOO_API_KEY (env vars, e.g.
via .env — same keys as BatteryAlerts.py).
"""

import os
import re
import sys
import csv
import argparse
import xmlrpc.client
from datetime import datetime

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


ODOO_URL      = os.getenv("ODOO_URL")
ODOO_DB       = os.getenv("ODOO_DB")
ODOO_USERNAME = os.getenv("ODOO_USERNAME")
ODOO_API_KEY  = os.getenv("ODOO_API_KEY")
ODOO_BATCH_SIZE = 200


class Odoo:
    def __init__(self):
        if not all([ODOO_URL, ODOO_DB, ODOO_USERNAME, ODOO_API_KEY]):
            sys.exit("Set ODOO_URL, ODOO_DB, ODOO_USERNAME and ODOO_API_KEY (e.g. in .env).")
        common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
        self.uid = common.authenticate(ODOO_DB, ODOO_USERNAME, ODOO_API_KEY, {})
        if not self.uid:
            sys.exit("Odoo authentication failed — check ODOO_DB/ODOO_USERNAME/ODOO_API_KEY.")
        self.models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object", allow_none=True)

    def search_read_all(self, model, domain, fields, order="id asc"):
        all_rows = []
        offset = 0
        while True:
            batch = self.models.execute_kw(
                ODOO_DB, self.uid, ODOO_API_KEY,
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


def find_sale_orders(odoo: Odoo, query: str) -> list:
    """Sale Orders whose Customer Reference contains every word of `query` (case-insensitive)."""
    words = query.split()
    domain = [["state", "!=", "cancel"]] + [["client_order_ref", "ilike", w] for w in words]
    return odoo.search_read_all(
        "sale.order", domain,
        ["id", "name", "client_order_ref", "partner_id"], order="id asc",
    )


def fetch_batteries(odoo: Odoo, orders: list) -> list:
    order_by_id = {o["id"]: o for o in orders}

    pickings = odoo.search_read_all(
        "stock.picking",
        [
            ["sale_id", "in", list(order_by_id)],
            ["picking_type_id.code", "=", "outgoing"],
            ["state", "=", "done"],
        ],
        ["id", "name", "sale_id"],
    )
    if not pickings:
        return []
    picking_by_id = {p["id"]: p for p in pickings}

    lines = odoo.search_read_all(
        "stock.move.line",
        [["picking_id", "in", list(picking_by_id)], ["lot_id", "!=", False]],
        ["lot_id", "product_id", "picking_id", "date"], order="date asc",
    )
    if not lines:
        return []

    # Current location per lot — highest-quantity quant wins, as in BatteryAlerts.
    lot_ids = list({ml["lot_id"][0] for ml in lines})
    quants = odoo.search_read_all(
        "stock.quant", [["lot_id", "in", lot_ids], ["quantity", ">", 0]],
        ["lot_id", "location_id", "quantity"], order="lot_id asc",
    )
    loc_ids = list({q["location_id"][0] for q in quants if q.get("location_id")})
    locs = odoo.search_read_all(
        "stock.location", [["id", "in", loc_ids]], ["id", "usage"],
    ) if loc_ids else []
    usage_by_loc = {loc["id"]: loc["usage"] for loc in locs}
    current_usage = {}
    best_qty = {}
    for q in quants:
        lot_id = q["lot_id"][0]
        if q["quantity"] > best_qty.get(lot_id, 0):
            best_qty[lot_id] = q["quantity"]
            current_usage[lot_id] = usage_by_loc.get(q["location_id"][0], "")

    # One row per (serial, sale order); lines are date-ascending, so a later
    # re-delivery on the same SO overwrites the earlier one.
    rows = {}
    for ml in lines:
        lot_id, serial = ml["lot_id"]
        picking = picking_by_id[ml["picking_id"][0]]
        order = order_by_id[picking["sale_id"][0]]
        rows[(serial, order["name"])] = {
            "Serial": serial,
            "Product": ml["product_id"][1] if ml.get("product_id") else "",
            "Sale Order": order["name"],
            "Customer Reference": order.get("client_order_ref") or "",
            "Invoiced Customer": order["partner_id"][1] if order.get("partner_id") else "",
            "Delivery": picking["name"],
            "Delivery Date": str(ml.get("date", ""))[:10],
            "Still At Customer": "Yes" if current_usage.get(lot_id) == "customer" else "No",
        }

    return sorted(rows.values(), key=lambda r: (r["Serial"], r["Delivery Date"]))


def main():
    parser = argparse.ArgumentParser(description="Export an end user's battery serials and Sale Orders from Odoo to CSV.")
    parser.add_argument("customer", nargs="?", help="End user name as written in the SO Customer Reference (any case, partial). Prompted for if omitted.")
    parser.add_argument("-o", "--output", help="Output CSV path (default: <name>_Batteries_<date>.csv)")
    args = parser.parse_args()

    query = (args.customer or input("Customer / end user name: ")).strip()
    if not query:
        sys.exit("No name entered.")

    odoo = Odoo()
    orders = find_sale_orders(odoo, query)
    if not orders:
        sys.exit(f"No Sale Order has a Customer Reference matching '{query}'.")

    print(f"\n{len(orders)} Sale Orders with '{query}' in the Customer Reference:")
    for o in orders:
        print(f"  {o['name']:<10} {o['client_order_ref']}")

    rows = fetch_batteries(odoo, orders)
    if not rows:
        sys.exit("\nNo delivered serials found on those Sale Orders.")

    safe_label = re.sub(r"[^A-Za-z0-9]+", "_", query.upper()).strip("_")
    output = args.output or f"{safe_label}_Batteries_{datetime.now():%Y-%m-%d}.csv"

    with open(output, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    serials = len({r["Serial"] for r in rows})
    so_count = len({r["Sale Order"] for r in rows})
    print(f"\nWrote {len(rows)} rows ({serials} serials across {so_count} sale orders) to {output}")


if __name__ == "__main__":
    main()
