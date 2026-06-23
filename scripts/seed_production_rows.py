#!/usr/bin/env python3
"""
Seed In Production rows in inventory.db for variants in the upcoming manifest.

For each target variant + quantity:
  1. get_or_create_variant in DB
  2. Insert qty In Production rows (blank serial)
  3. Regenerate Excel

Run this before parse_manifest.py to simulate units ordered and put into production.

Usage:
    python3 scripts/seed_production_rows.py
"""

import os
import sys

_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SCRIPTS_DIR)

from constants import get_sku
from db import get_conn, get_or_create_variant, add_unit
from export_excel import export_excel

# Variants to seed and how many In Production rows to add.
# (Design Name, Size, Finish, Swing, Glass Type): qty
SEED = {
    ("Valencia Single Door", '32" x 80"', "Matte Black", "Left Inswing", "Clear"): 3,
    ("Sevilla French Door",  '36" x 80"', "Satin White", "Left Inswing", "None"):  2,
    ("Altea Double Door",    '36" x 80"', "Matte Black", "Left Inswing", "Rain"):  2,
}


def main():
    conn = get_conn()

    for variant, qty in SEED.items():
        design, size, finish, swing, glass = variant
        sku = get_sku(variant)
        vid = get_or_create_variant(conn, design, size, finish, swing, glass, sku)

        for _ in range(qty):
            add_unit(conn, vid, None, "In Production")

        print(f"  Seeded {qty} In Production row(s) for {design} {size} {finish}")

    print("\nRegenerating Excel ...")
    export_excel(conn)
    conn.close()
    print("Done.")


if __name__ == "__main__":
    main()
