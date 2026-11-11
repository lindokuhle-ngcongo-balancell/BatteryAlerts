#!/usr/bin/env python3
"""List batteries delivered on Odoo sales orders related to a contact.

Usage:
    python batteries_for_contact.py
    python batteries_for_contact.py "Marco Carelli"

The search matches every query word in Odoo contact names, then finds non-cancelled sales
orders for those contacts, their commercial partner and its child contacts.
Sales orders whose Customer Reference contains every query word are also included.
Battery details come from done outgoing deliveries on those sales orders.
Requires the ODOO_URL, ODOO_DB, ODOO_USERNAME and ODOO_API_KEY environment
variables (for example, in .env).
"""

import argparse
import csv
import sys

from customer_batteries import Odoo, fetch_batteries


def matches_query(value: str, query: str) -> bool:
    normalized_value = value.casefold()
    return all(word.casefold() in normalized_value for word in query.split())


def find_contacts(odoo: Odoo, query: str) -> list:
    domain = [["name", "ilike", word] for word in query.split()]
    contacts = odoo.search_read_all(
        "res.partner",
        domain,
        ["id", "display_name", "commercial_partner_id"],
        order="id asc",
    )
    return [
        contact for contact in contacts
        if matches_query(contact.get("display_name", ""), query)
    ]


def find_related_orders(odoo: Odoo, query: str, contacts: list) -> list:
    contact_ids = [contact["id"] for contact in contacts]
    reference_domain = [
        ["client_order_ref", "ilike", word] for word in query.split()
    ]
    company_ids = list({
        contact["commercial_partner_id"][0]
        for contact in contacts
        if contact.get("commercial_partner_id")
    })

    domain = [["state", "!=", "cancel"]]
    if contact_ids and company_ids:
        domain += ["|", "|", ["partner_id", "in", contact_ids]]
        domain += [["partner_id", "child_of", company_ids]]
        domain += ["&"] * (len(reference_domain) - 1) + reference_domain
    elif contact_ids:
        domain += ["|", ["partner_id", "in", contact_ids]]
        domain += ["&"] * (len(reference_domain) - 1) + reference_domain
    else:
        domain += reference_domain

    return odoo.search_read_all(
        "sale.order",
        domain,
        ["id", "name", "client_order_ref", "partner_id"],
        order="id asc",
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Find batteries delivered on Odoo sales orders related to a contact."
    )
    parser.add_argument("contact", nargs="?", default="Marco Carelli", help="Contact name to search for.")
    args = parser.parse_args()
    query = args.contact.strip().casefold()
    if not query:
        parser.error("contact name must not be empty")

    odoo = Odoo()
    contacts = find_contacts(odoo, query)
    orders = find_related_orders(odoo, query, contacts)

    print(f"Contact matches for {query!r}: {len(contacts)}")
    for contact in contacts:
        print(f"  {contact['display_name']}")

    print(f"\nRelated non-cancelled sales orders: {len(orders)}")
    for order in orders:
        print(f"  {order['name']}: {order.get('client_order_ref') or ''}")

    batteries = fetch_batteries(odoo, orders)
    print(f"\nBatteries found: {len(batteries)}")
    if batteries:
        writer = csv.DictWriter(
            sys.stdout,
            fieldnames=list(batteries[0]),
            delimiter="\t",
            lineterminator="\n",
        )
        writer.writeheader()
        writer.writerows(batteries)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
