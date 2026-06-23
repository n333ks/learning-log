#!/usr/bin/env python3
"""
Parse a container manifest CSV and update inventory.db (then regenerate Excel).

For each manifest line item:
  MATCHED variant  → fill In Production rows (blank serial) with arriving serials,
                     flip status to "Pre-Sale - <container_id>";
                     excess serials → add_unit with Pre-Sale status
  NEW variant      → get_or_create_variant, then add_unit for each serial
  Has Customer Order → get_or_create_variant, then add_sales_order for each serial

Usage:
    python3 scripts/parse_manifest.py --container CNT-001
    python3 scripts/parse_manifest.py --container CNT-001 --manifest container_manifest.csv
"""

import argparse
import csv
import glob
import os
import sys

_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SCRIPTS_DIR)

from constants import get_sku, MANIFEST_PATTERN
from db import (
    get_conn, get_variant_id, get_or_create_variant,
    fill_production_unit, count_production_units, add_unit, add_sales_order,
)
from export_excel import export_excel

VARIANT_KEYS = ["Design Name", "Size", "Finish", "Swing", "Glass Type"]


# ── Manifest helpers ───────────────────────────────────────────────────────────

def load_manifest(path):
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def manifest_key(row):
    return tuple(str(row[k] or "").strip() for k in VARIANT_KEYS)


def manifest_serials(row):
    val = row.get("Serial Numbers") or ""
    return [s.strip() for s in str(val).split(";") if s.strip()]


def manifest_status(row, container_id):
    """Return Allocated status if a customer order exists, otherwise Pre-Sale."""
    order = str(row.get("Customer Order") or "").strip()
    if order:
        return f"Allocated - {order}"
    return f"Pre-Sale - {container_id}"


def pick_manifest():
    """Find all container manifests and let the user choose one interactively."""
    files = sorted(glob.glob(MANIFEST_PATTERN))
    if not files:
        print(f"No manifest files found matching '{MANIFEST_PATTERN}'.")
        raise SystemExit(1)

    print("Available manifests:")
    for i, f in enumerate(files, 1):
        print(f"  [{i}] {f}")

    while True:
        choice = input("\nSelect a manifest to parse (number): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(files):
            return files[int(choice) - 1]
        print(f"  Please enter a number between 1 and {len(files)}.")


def container_id_from_filename(path):
    """Extract CNT-XXX from container_manifest_CNT-XXX.csv."""
    base = os.path.basename(path)
    name = os.path.splitext(base)[0]
    return name.split("container_manifest_")[-1]


# ── Core logic ─────────────────────────────────────────────────────────────────

def process_manifest(conn, manifest_rows, container_id):
    """
    Process all manifest rows against the DB.
    Returns (n_matched, n_new) counts.
    """
    pre_sale_status = f"Pre-Sale - {container_id}"
    n_matched = 0
    n_new     = 0

    for item in manifest_rows:
        key            = manifest_key(item)
        serials        = manifest_serials(item)
        customer_order = str(item.get("Customer Order") or "").strip()
        design, size, finish, swing, glass = key
        sku = get_sku(key)

        if customer_order:
            # ── Allocated → sales_orders ──────────────────────────────────────
            # Parse "ORD-XXXX - Customer Name"
            if " - " in customer_order:
                order_num, customer = customer_order.split(" - ", 1)
            else:
                order_num = customer_order
                customer  = ""

            vid = get_or_create_variant(conn, design, size, finish, swing, glass, sku)
            for serial in serials:
                add_sales_order(conn, order_num, customer, vid, serial, container_id)
            print(f"  SALES       {design} {size} {finish} — "
                  f"{len(serials)} serial(s) → Sales ({customer_order})")

        else:
            vid = get_variant_id(conn, design, size, finish, swing, glass)

            if vid is not None:
                # ── Matched: fill In Production rows, queue excess ──────────
                n_matched += 1
                prod_count = count_production_units(conn, vid)
                to_fill    = min(prod_count, len(serials))
                excess     = serials[to_fill:]

                for serial in serials[:to_fill]:
                    fill_production_unit(conn, vid, serial, pre_sale_status, container_id)

                for serial in excess:
                    add_unit(conn, vid, serial, pre_sale_status, container_id)

                if excess:
                    print(f"  MATCHED     {design} {size} {finish} — "
                          f"{to_fill} In Production filled + {len(excess)} excess → Pre-Sale")
                elif to_fill:
                    print(f"  MATCHED     {design} {size} {finish} — "
                          f"{to_fill} serial(s) → {pre_sale_status}")
                else:
                    # No In Production rows at all — all go straight to Pre-Sale
                    for serial in serials:
                        add_unit(conn, vid, serial, pre_sale_status, container_id)
                    print(f"  MATCHED     {design} {size} {finish} — "
                          f"{len(serials)} serial(s) → {pre_sale_status} (no In Production rows)")

            else:
                # ── New variant ───────────────────────────────────────────────
                n_new += 1
                vid = get_or_create_variant(conn, design, size, finish, swing, glass, sku)
                for serial in serials:
                    add_unit(conn, vid, serial, pre_sale_status, container_id)
                print(f"  NEW VARIANT {design} {size} {finish} — "
                      f"{len(serials)} row(s) → {pre_sale_status}")

    return n_matched, n_new


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Parse a container manifest and update inventory."
    )
    parser.add_argument("--manifest",  default=None,
                        help="Manifest CSV to parse (default: interactive selection)")
    parser.add_argument("--container", default=None,
                        help="Container ID override (default: derived from filename)")
    args = parser.parse_args()

    manifest_path = args.manifest or pick_manifest()
    container_id  = args.container or container_id_from_filename(manifest_path)

    print(f"\nLoading manifest: {manifest_path}  (container: {container_id})")
    manifest_rows = load_manifest(manifest_path)
    print(f"  {len(manifest_rows)} line item(s)\n")

    conn = get_conn()
    n_matched, n_new = process_manifest(conn, manifest_rows, container_id)

    print(f"\nRegenerating Excel workbook ...")
    export_excel(conn)
    conn.close()

    print(f"\nSummary: {n_matched} matched variant(s), {n_new} new variant(s) created.")


if __name__ == "__main__":
    main()
