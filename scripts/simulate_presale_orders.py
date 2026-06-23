"""
Simulate 5-10 Shopify pre-sale orders per container (CNT-002 through CNT-005),
interleaved as if all containers were available simultaneously.
Then renumber ALL orders sequentially: ORD-00001, ORD-00002, ...
"""

import sqlite3, random
from pathlib import Path

DB = Path(__file__).parent.parent / "inventory.db"

# Orders to create: (container, unit_ids, customer, date)
# Mixed across containers in realistic date order (late May – mid June 2026)
ORDERS = [
    # date, container, unit_ids (one per unit in the order), customer
    ("2026-05-28", "CNT-004", [234], "Alex Hernandez"),
    ("2026-05-29", "CNT-002", [228], "Taylor Morrison"),
    ("2026-05-30", "CNT-003", [14], "Brandon Walsh"),
    ("2026-05-31", "CNT-005", [241], "Sofia Rodriguez"),
    ("2026-06-01", "CNT-004", [142], "Kevin Park"),
    ("2026-06-02", "CNT-002", [135, 136], "Lisa Chen"),
    ("2026-06-03", "CNT-003", [83], "Marcus Williams"),
    ("2026-06-04", "CNT-005", [268], "Stephanie Davis"),
    ("2026-06-05", "CNT-004", [98], "Ryan Thompson"),
    ("2026-06-06", "CNT-002", [32], "Amanda Garcia"),
    ("2026-06-07", "CNT-003", [96], "Daniel Lee"),
    ("2026-06-08", "CNT-005", [277, 278], "Rachel Kim"),
    ("2026-06-09", "CNT-004", [200], "Carlos Mendez"),
    ("2026-06-10", "CNT-002", [24], "Hannah Brown"),
    ("2026-06-11", "CNT-003", [118], "Tyler Johnson"),
    ("2026-06-12", "CNT-005", [261], "Ashley Martinez"),
    ("2026-06-13", "CNT-004", [30, 31], "Derek Wilson"),
    ("2026-06-14", "CNT-002", [153], "Nicole Baker"),
    ("2026-06-15", "CNT-003", [38, 39], "Justin Nguyen"),
    ("2026-06-16", "CNT-005", [253, 254], "Megan Foster"),
    ("2026-06-17", "CNT-004", [162], "Connor Hayes"),
    ("2026-06-18", "CNT-002", [188], "Lauren Phillips"),
    ("2026-06-19", "CNT-003", [150], "Ethan Cruz"),
    ("2026-06-20", "CNT-005", [259, 260], "Brittany Sanders"),
    ("2026-06-21", "CNT-004", [186, 187], "Jacob Reed"),
    ("2026-06-21", "CNT-002", [40, 41], "Olivia Bell"),
    ("2026-06-22", "CNT-003", [169], "Nathan Castillo"),
    ("2026-06-22", "CNT-005", [257], "Emma Clarke"),
]

def run():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")

    # ── Step 1: insert new pre-sale orders ───────────────────────────────────
    # Use placeholder order numbers first; we'll renumber everything after
    new_order_rows = []
    placeholder_start = 9000  # well out of range of existing numbers

    for i, (date, container, unit_ids, customer) in enumerate(ORDERS):
        placeholder = f"ORD-{placeholder_start + i}"
        # Get variant_id from the first unit
        row = conn.execute("SELECT variant_id FROM units WHERE id = ?", (unit_ids[0],)).fetchone()
        if row is None:
            print(f"  SKIP: unit {unit_ids[0]} not found")
            continue
        variant_id = row["variant_id"]

        # Insert one sales_order row per unit_id
        for uid in unit_ids:
            unit = conn.execute("SELECT status, variant_id FROM units WHERE id = ?", (uid,)).fetchone()
            if unit is None:
                print(f"  SKIP: unit {uid} not found")
                continue
            if "Pre-Sale" not in (unit["status"] or ""):
                print(f"  SKIP: unit {uid} is '{unit['status']}', not Pre-Sale")
                continue

            conn.execute("""
                INSERT INTO sales_orders (order_number, customer, variant_id,
                    serial_number, date_allocated, status, container_id)
                SELECT ?, ?, variant_id, serial_number, ?, 'Allocated', container_id
                FROM units WHERE id = ?
            """, (placeholder, customer, date, uid))

            conn.execute("""
                UPDATE units SET status = ? WHERE id = ?
            """, (f"Allocated ({placeholder})", uid))

        new_order_rows.append(placeholder)
        print(f"  Created {placeholder}: {customer} ({container}, {len(unit_ids)} unit(s))")

    conn.commit()

    # ── Step 2: renumber ALL orders sequentially ─────────────────────────────
    # Collect all unique order numbers across both tables, ordered by earliest date
    all_orders = conn.execute("""
        SELECT order_number, MIN(date_allocated) AS earliest
        FROM (
            SELECT order_number, date_allocated FROM sales_orders
            UNION ALL
            SELECT order_number, date_arrived AS date_allocated FROM warehouse
        )
        GROUP BY order_number
        ORDER BY earliest, order_number
    """).fetchall()

    renumber_map = {}
    for seq, row in enumerate(all_orders, start=1):
        renumber_map[row["order_number"]] = f"ORD-{seq:05d}"

    print("\nRenumbering map:")
    for old, new in renumber_map.items():
        print(f"  {old} → {new}")

    # Apply to sales_orders
    for old, new in renumber_map.items():
        conn.execute("UPDATE sales_orders SET order_number = ? WHERE order_number = ?", (new, old))
        # Also fix the status field in units (e.g. "Allocated (ORD-9000)")
        conn.execute(
            "UPDATE units SET status = ? WHERE status = ?",
            (f"Allocated ({new})", f"Allocated ({old})")
        )

    # Apply to warehouse
    for old, new in renumber_map.items():
        conn.execute("UPDATE warehouse SET order_number = ? WHERE order_number = ?", (new, old))

    conn.commit()
    conn.close()
    print(f"\nDone. {len(ORDERS)} new orders created, {len(renumber_map)} total orders renumbered.")

if __name__ == "__main__":
    run()
