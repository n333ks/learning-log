#!/usr/bin/env python3
"""
Move rows from the Sales table to the Warehouse table when a container arrives,
or cancel an order and return units to the Inventory table.

Usage:
    python3 scripts/receive_container.py
    python3 scripts/receive_container.py --cancel ORD-1001
"""

import argparse
import datetime
import os
import sys

_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SCRIPTS_DIR)

from db import (
    get_conn, get_all_containers_in_sales, get_sales_by_container,
    move_to_warehouse, cancel_order as db_cancel_order, add_unit,
)
from export_excel import export_excel


# ── Container arrival flow ─────────────────────────────────────────────────────

def container_arrival(conn):
    containers = get_all_containers_in_sales(conn)

    if not containers:
        print("No allocated units found in Sales.")
        return

    containers = sorted(containers)
    print("Containers with allocated units:")
    for i, cid in enumerate(containers, 1):
        count = len(get_sales_by_container(conn, cid))
        print(f"  [{i}] {cid}  ({count} unit{'s' if count != 1 else ''})")

    while True:
        choice = input("\nSelect a container (number): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(containers):
            break
        print(f"  Please enter a number between 1 and {len(containers)}.")

    selected = containers[int(choice) - 1]
    sales    = get_sales_by_container(conn, selected)
    today    = datetime.date.today().isoformat()

    print(f"\nMoving {len(sales)} unit(s) from Sales → Warehouse (Date Arrived: {today}) ...")
    for sale in sales:
        move_to_warehouse(conn, sale["id"], today)
        print(f"  {sale['serial_number']}  [{sale['order_number']}]")

    print(f"\nMoved {len(sales)} unit(s). Regenerating Excel ...")
    export_excel(conn)


# ── Cancellation flow ──────────────────────────────────────────────────────────

def cancel_order_flow(conn, order_num):
    units = db_cancel_order(conn, order_num)

    if not units:
        print(f"No Sales rows found for order {order_num}.")
        return

    print(f"Cancelling {order_num}: {len(units)} unit(s) found.")
    for variant_id, serial, container_id in units:
        status = f"Pre-Sale - {container_id}" if container_id else "Pre-Sale"
        add_unit(conn, variant_id, serial, status, container_id)
        print(f"  Returned {serial} → Inventory as {status}")

    print("\nRegenerating Excel ...")
    export_excel(conn)


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Receive a container (Sales → Warehouse) or cancel an order."
    )
    parser.add_argument("--cancel", metavar="ORDER_NUM", default=None,
                        help="Cancel an order and return units to Inventory")
    args = parser.parse_args()

    conn = get_conn()

    if args.cancel:
        cancel_order_flow(conn, args.cancel)
    else:
        container_arrival(conn)

    conn.close()
    print("\nDone.")


if __name__ == "__main__":
    main()
