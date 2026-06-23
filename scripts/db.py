#!/usr/bin/env python3
"""
Database layer for the inventory management system.
All scripts import from here instead of using openpyxl directly.
"""

import sqlite3
import os

from constants import _BASE_DIR

DB_PATH = os.path.join(_BASE_DIR, "inventory.db")

CHECKLIST_ITEMS = [
    ("serial_verified",   "Serial number verified against order"),
    ("design_confirmed",  "Design / model confirmed"),
    ("size_confirmed",    "Dimensions confirmed (W × H)"),
    ("finish_confirmed",  "Finish / color matches order"),
    ("glass_type",        "Glass type confirmed"),
    ("glass_condition",   "Glass panels intact — no cracks or chips"),
    ("frame_condition",   "Frame inspected — no dents, scratches, or weld defects"),
    ("finish_quality",    "Finish quality — uniform coating, no bare spots or runs"),
    ("hardware_package",  "Hardware package included (hinges, handles, locks, keys)"),
    ("weatherstrip",      "Weatherstripping attached and intact"),
    ("mounting_hardware", "All mounting / installation hardware included"),
    ("cleaned_wrapped",   "Product cleaned and wrapped for delivery"),
    ("crating_complete",  "Crating / packaging complete and secure"),
    ("photos_taken",      "Final product photos uploaded"),
]

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    username      TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role          TEXT NOT NULL DEFAULT 'warehouse',
    full_name     TEXT
);

CREATE TABLE IF NOT EXISTS variants (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    design_name   TEXT NOT NULL,
    size          TEXT NOT NULL,
    finish        TEXT NOT NULL,
    swing         TEXT NOT NULL,
    glass_type    TEXT NOT NULL,
    sku           TEXT NOT NULL,
    optimal_count INTEGER DEFAULT 0,
    UNIQUE(design_name, size, finish, swing, glass_type)
);

CREATE TABLE IF NOT EXISTS units (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    variant_id    INTEGER NOT NULL REFERENCES variants(id),
    serial_number TEXT,
    status        TEXT NOT NULL,
    container_id  TEXT,
    date_added    TEXT DEFAULT (date('now'))
);

CREATE TABLE IF NOT EXISTS sales_orders (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    order_number     TEXT NOT NULL,
    customer         TEXT NOT NULL,
    variant_id       INTEGER NOT NULL REFERENCES variants(id),
    serial_number    TEXT NOT NULL,
    container_id     TEXT,
    date_allocated   TEXT DEFAULT (date('now')),
    status           TEXT DEFAULT 'Allocated',
    fulfillment_type TEXT DEFAULT 'pickup'
);

CREATE TABLE IF NOT EXISTS warehouse (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    order_number     TEXT NOT NULL,
    customer         TEXT NOT NULL,
    variant_id       INTEGER NOT NULL REFERENCES variants(id),
    serial_number    TEXT NOT NULL,
    container_id     TEXT,
    date_arrived     TEXT DEFAULT (date('now')),
    status           TEXT DEFAULT 'In Prep',
    notes            TEXT,
    fulfillment_type TEXT DEFAULT 'pickup'
);

CREATE TABLE IF NOT EXISTS warehouse_checklist (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    warehouse_id INTEGER NOT NULL REFERENCES warehouse(id) ON DELETE CASCADE,
    item_key     TEXT NOT NULL,
    label        TEXT NOT NULL,
    completed    INTEGER DEFAULT 0,
    completed_at TEXT,
    completed_by TEXT,
    UNIQUE(warehouse_id, item_key)
);

CREATE TABLE IF NOT EXISTS warehouse_photos (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    warehouse_id INTEGER NOT NULL REFERENCES warehouse(id) ON DELETE CASCADE,
    filename     TEXT NOT NULL,
    caption      TEXT,
    uploaded_at  TEXT,
    uploaded_by  TEXT
);

CREATE TABLE IF NOT EXISTS activity_log (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    username      TEXT,
    full_name     TEXT,
    action        TEXT NOT NULL,
    detail        TEXT,
    order_number  TEXT,
    serial_number TEXT,
    warehouse_id  INTEGER,
    created_at    TEXT
);

CREATE TABLE IF NOT EXISTS change_requests (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    order_number   TEXT NOT NULL,
    customer       TEXT NOT NULL,
    order_type     TEXT NOT NULL,
    scenario_id    TEXT NOT NULL,
    request_detail TEXT,
    notes          TEXT,
    status         TEXT DEFAULT 'Open',
    resolution     TEXT,
    created_by     TEXT,
    created_at     TEXT,
    updated_at     TEXT
);
"""


def get_conn():
    """Return a sqlite3 connection with row_factory=sqlite3.Row."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    """Create all tables if they do not exist, and run any pending migrations."""
    conn = get_conn()
    conn.executescript(SCHEMA)
    for sql in [
        "ALTER TABLE warehouse ADD COLUMN notes TEXT",
        "ALTER TABLE activity_log ADD COLUMN warehouse_id INTEGER",
        "ALTER TABLE warehouse ADD COLUMN fulfillment_type TEXT DEFAULT 'pickup'",
        "ALTER TABLE sales_orders ADD COLUMN fulfillment_type TEXT DEFAULT 'pickup'",
    ]:
        try:
            conn.execute(sql)
        except Exception:
            pass
    conn.commit()
    conn.close()


# ── Activity log ──────────────────────────────────────────────────────────────

def log_activity(conn, username, full_name, action, detail=None, order_number=None, serial_number=None, warehouse_id=None):
    from datetime import datetime
    conn.execute(
        "INSERT INTO activity_log (username, full_name, action, detail, order_number, serial_number, warehouse_id, created_at) VALUES (?,?,?,?,?,?,?,?)",
        (username, full_name, action, detail, order_number, serial_number, warehouse_id, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    )
    conn.commit()


def get_activity_log(conn, limit=300):
    return conn.execute(
        "SELECT * FROM activity_log ORDER BY id DESC LIMIT ?", (limit,)
    ).fetchall()


# ── User management ────────────────────────────────────────────────────────────

def get_user(conn, username):
    return conn.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()


def get_all_users(conn):
    return conn.execute("SELECT id, username, role, full_name FROM users ORDER BY role, username").fetchall()


def create_user(conn, username, password_hash, role, full_name=""):
    conn.execute(
        "INSERT INTO users (username, password_hash, role, full_name) VALUES (?,?,?,?)",
        (username, password_hash, role, full_name),
    )
    conn.commit()


def update_user_password(conn, user_id, password_hash):
    conn.execute("UPDATE users SET password_hash=? WHERE id=?", (password_hash, user_id))
    conn.commit()


def delete_user(conn, user_id):
    conn.execute("DELETE FROM users WHERE id=?", (user_id,))
    conn.commit()


def seed_admin(conn, username, password_hash):
    """Create the admin user if no users exist yet."""
    count = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
    if count == 0:
        conn.execute(
            "INSERT INTO users (username, password_hash, role, full_name) VALUES (?,?,'admin','Administrator')",
            (username, password_hash),
        )
        conn.commit()


# ── Warehouse prep ─────────────────────────────────────────────────────────────

def ensure_checklist(conn, warehouse_id):
    """Create checklist rows for a single warehouse unit if they don't exist yet."""
    existing = conn.execute(
        "SELECT COUNT(*) FROM warehouse_checklist WHERE warehouse_id=?", (warehouse_id,)
    ).fetchone()[0]
    if existing == 0:
        conn.executemany(
            "INSERT OR IGNORE INTO warehouse_checklist (warehouse_id, item_key, label) VALUES (?,?,?)",
            [(warehouse_id, key, label) for key, label in CHECKLIST_ITEMS],
        )
        conn.commit()


def toggle_checklist_item(conn, warehouse_id, item_key, completed, username):
    now = __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M")
    conn.execute("""
        UPDATE warehouse_checklist
        SET completed=?, completed_at=?, completed_by=?
        WHERE warehouse_id=? AND item_key=?
    """, (1 if completed else 0, now if completed else None, username if completed else None,
          warehouse_id, item_key))
    conn.commit()


def save_warehouse_notes(conn, warehouse_id, notes):
    conn.execute("UPDATE warehouse SET notes=? WHERE id=?", (notes, warehouse_id))
    conn.commit()


def add_warehouse_photo(conn, warehouse_id, filename, caption, username):
    now = __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M")
    conn.execute(
        "INSERT INTO warehouse_photos (warehouse_id, filename, caption, uploaded_at, uploaded_by) VALUES (?,?,?,?,?)",
        (warehouse_id, filename, caption, now, username),
    )
    conn.commit()


def delete_warehouse_photo(conn, photo_id):
    row = conn.execute("SELECT filename, warehouse_id FROM warehouse_photos WHERE id=?", (photo_id,)).fetchone()
    conn.execute("DELETE FROM warehouse_photos WHERE id=?", (photo_id,))
    conn.commit()
    return (row["filename"], row["warehouse_id"]) if row else (None, None)


def mark_warehouse_ready(conn, warehouse_id):
    conn.execute("UPDATE warehouse SET status='Pending Review' WHERE id=?", (warehouse_id,))
    conn.commit()


def approve_warehouse_order(conn, order_number):
    """Accept a Pending Review order — sets each unit to Ready for Pickup or Ready for Delivery."""
    conn.execute("""
        UPDATE warehouse
        SET status = CASE
            WHEN fulfillment_type='delivery' THEN 'Ready for Delivery'
            ELSE 'Ready for Pickup'
        END
        WHERE order_number=?
    """, (order_number,))
    conn.commit()


def reject_warehouse_order(conn, order_number):
    """Push back a Pending Review order — resets all units to In Prep."""
    conn.execute("UPDATE warehouse SET status='In Prep' WHERE order_number=?", (order_number,))
    conn.commit()


def get_prep_data(conn, warehouse_id):
    """Return (unit row, checklist rows, photos) for a single warehouse unit."""
    unit = conn.execute("""
        SELECT wh.id, wh.order_number, wh.customer, wh.serial_number, wh.container_id,
               wh.date_arrived, wh.status, wh.notes, wh.fulfillment_type,
               v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku
        FROM warehouse wh JOIN variants v ON v.id = wh.variant_id
        WHERE wh.id=?
    """, (warehouse_id,)).fetchone()
    checklist = conn.execute(
        "SELECT * FROM warehouse_checklist WHERE warehouse_id=? ORDER BY id", (warehouse_id,)
    ).fetchall()
    photos = conn.execute(
        "SELECT * FROM warehouse_photos WHERE warehouse_id=? ORDER BY id", (warehouse_id,)
    ).fetchall()
    return unit, checklist, photos


def get_order_prep_summary(conn, order_number):
    """Return (units_with_progress, all_photos) for an order summary page."""
    units = conn.execute("""
        SELECT wh.id, wh.order_number, wh.customer, wh.serial_number, wh.container_id,
               wh.date_arrived, wh.status, wh.notes, wh.fulfillment_type,
               v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku
        FROM warehouse wh JOIN variants v ON v.id = wh.variant_id
        WHERE wh.order_number=?
        ORDER BY wh.id
    """, (order_number,)).fetchall()

    unit_ids = [u['id'] for u in units]
    total_items = len(CHECKLIST_ITEMS)

    progress = {}
    if unit_ids:
        placeholders = ','.join('?' * len(unit_ids))
        rows = conn.execute(
            f"SELECT warehouse_id, COUNT(*) AS done FROM warehouse_checklist WHERE warehouse_id IN ({placeholders}) AND completed=1 GROUP BY warehouse_id",
            unit_ids
        ).fetchall()
        progress = {r['warehouse_id']: r['done'] for r in rows}

    photos = []
    if unit_ids:
        placeholders = ','.join('?' * len(unit_ids))
        photos = conn.execute(
            f"SELECT wp.*, wh.serial_number FROM warehouse_photos wp JOIN warehouse wh ON wh.id = wp.warehouse_id WHERE wp.warehouse_id IN ({placeholders}) ORDER BY wp.id",
            unit_ids
        ).fetchall()

    return units, progress, total_items, photos


def get_or_create_variant(conn, design, size, finish, swing, glass, sku):
    """Return variant id. Inserts row if the variant does not exist yet."""
    row = conn.execute(
        "SELECT id FROM variants WHERE design_name=? AND size=? AND finish=? AND swing=? AND glass_type=?",
        (design, size, finish, swing, glass),
    ).fetchone()
    if row:
        return row["id"]
    cur = conn.execute(
        "INSERT INTO variants (design_name, size, finish, swing, glass_type, sku) VALUES (?,?,?,?,?,?)",
        (design, size, finish, swing, glass, sku),
    )
    conn.commit()
    return cur.lastrowid


def get_variant_id(conn, design, size, finish, swing, glass):
    """Return variant id or None if not found."""
    row = conn.execute(
        "SELECT id FROM variants WHERE design_name=? AND size=? AND finish=? AND swing=? AND glass_type=?",
        (design, size, finish, swing, glass),
    ).fetchone()
    return row["id"] if row else None


def add_unit(conn, variant_id, serial, status, container_id=None):
    """Insert a unit row and commit."""
    conn.execute(
        "INSERT INTO units (variant_id, serial_number, status, container_id) VALUES (?,?,?,?)",
        (variant_id, serial, status, container_id),
    )
    conn.commit()


def fill_production_unit(conn, variant_id, serial, status, container_id=None):
    """
    Find one In Production unit (no serial) for this variant and set its
    serial_number, status, and container_id. Returns True if updated, False if none found.
    """
    row = conn.execute(
        "SELECT id FROM units WHERE variant_id=? AND status='In Production' AND (serial_number IS NULL OR serial_number='') LIMIT 1",
        (variant_id,),
    ).fetchone()
    if not row:
        return False
    conn.execute(
        "UPDATE units SET serial_number=?, status=?, container_id=? WHERE id=?",
        (serial, status, container_id, row["id"]),
    )
    conn.commit()
    return True


def count_production_units(conn, variant_id):
    """Return count of In Production rows with no serial for this variant."""
    row = conn.execute(
        "SELECT COUNT(*) AS n FROM units WHERE variant_id=? AND status='In Production' AND (serial_number IS NULL OR serial_number='')",
        (variant_id,),
    ).fetchone()
    return row["n"]


def get_units_summary(conn):
    """
    Return rows with variant fields + in_stock, in_prod, pre_sale,
    optimal_count, variance — one row per variant.
    """
    rows = conn.execute("""
        SELECT
            v.id,
            v.design_name,
            v.size,
            v.finish,
            v.swing,
            v.glass_type,
            v.sku,
            v.optimal_count,
            SUM(CASE WHEN u.status = 'In Stock' THEN 1 ELSE 0 END)          AS in_stock,
            SUM(CASE WHEN u.status = 'In Production' THEN 1 ELSE 0 END)     AS in_prod,
            SUM(CASE WHEN u.status LIKE 'Pre-Sale%' THEN 1 ELSE 0 END)      AS pre_sale,
            SUM(CASE WHEN u.status = 'In Stock' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN u.status = 'In Production' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN u.status LIKE 'Pre-Sale%' THEN 1 ELSE 0 END)
              - v.optimal_count                                               AS variance
        FROM variants v
        LEFT JOIN units u ON u.variant_id = v.id
        GROUP BY v.id
        ORDER BY v.design_name, v.size, v.finish, v.swing, v.glass_type
    """).fetchall()
    return rows


def add_sales_order(conn, order_number, customer, variant_id, serial, container_id):
    """Insert a sales_order row and commit."""
    conn.execute(
        "INSERT INTO sales_orders (order_number, customer, variant_id, serial_number, container_id) VALUES (?,?,?,?,?)",
        (order_number, customer, variant_id, serial, container_id),
    )
    conn.commit()


def get_sales_by_container(conn, container_id):
    """Return all sales_orders rows for that container with status='Allocated'."""
    return conn.execute(
        "SELECT * FROM sales_orders WHERE container_id=? AND status='Allocated'",
        (container_id,),
    ).fetchall()


def move_to_warehouse(conn, sale_id, date_arrived):
    """Copy a sales_order row to warehouse, then delete from sales_orders."""
    sale = conn.execute("SELECT * FROM sales_orders WHERE id=?", (sale_id,)).fetchone()
    if not sale:
        return
    ft = sale["fulfillment_type"] or 'pickup'
    conn.execute(
        "INSERT INTO warehouse (order_number, customer, variant_id, serial_number, container_id, date_arrived, status, fulfillment_type) VALUES (?,?,?,?,?,?,'In Prep',?)",
        (sale["order_number"], sale["customer"], sale["variant_id"],
         sale["serial_number"], sale["container_id"], date_arrived, ft),
    )
    conn.execute("DELETE FROM sales_orders WHERE id=?", (sale_id,))
    conn.commit()


def _restore_status(conn, serial, container_id):
    """Return the correct status when releasing a unit back to inventory."""
    row = conn.execute("SELECT date_received FROM units WHERE serial_number=?", (serial,)).fetchone()
    if row and row["date_received"]:
        return "In Stock"  # container already arrived
    return f"Pre-Sale ({container_id})" if container_id else "In Stock"


def cancel_order(conn, order_number):
    """
    Cancel all sales_orders rows for order_number.
    Returns each unit to its correct status and creates a new In Production unit
    for each cancelled variant so FIFO stock replacement is queued automatically.
    """
    rows = conn.execute(
        "SELECT variant_id, serial_number, container_id FROM sales_orders WHERE order_number=?",
        (order_number,),
    ).fetchall()
    for r in rows:
        restore = _restore_status(conn, r["serial_number"], r["container_id"])
        conn.execute("UPDATE units SET status=? WHERE serial_number=?", (restore, r["serial_number"]))
    conn.execute("DELETE FROM sales_orders WHERE order_number=?", (order_number,))
    conn.commit()
    return len(rows)


def cancel_warehouse_order(conn, order_number, serial):
    """
    Cancel one warehouse row. Unit returns to In Stock and a replacement
    In Production unit is queued for that variant.
    """
    conn.execute("UPDATE units SET status='In Stock' WHERE serial_number=?", (serial,))
    conn.execute("DELETE FROM warehouse WHERE order_number=? AND serial_number=?", (order_number, serial))
    conn.commit()


def change_order_unit(conn, order_number, old_serial, new_unit_id, source="sales"):
    """
    Swap one unit on an existing order (works for both sales_orders and warehouse rows).
    - source: 'sales' or 'warehouse'
    - old_serial: serial being replaced
    - new_unit_id: units.id of the replacement (must be In Stock or Pre-Sale)
    """
    # Validate new unit
    new_unit = conn.execute(
        "SELECT id, variant_id, serial_number, status, container_id FROM units WHERE id=?",
        (new_unit_id,),
    ).fetchone()
    if not new_unit:
        raise ValueError(f"Unit {new_unit_id} not found")
    if new_unit["status"] != "In Stock" and not new_unit["status"].startswith("Pre-Sale"):
        raise ValueError(f"Unit {new_unit_id} is not available (status: {new_unit['status']})")

    # Release old unit back to inventory
    if source == "warehouse":
        # Container has physically arrived (unit is in warehouse), so always In Stock
        restore = "In Stock"
    else:
        old_unit = conn.execute(
            "SELECT container_id FROM units WHERE serial_number=?", (old_serial,)
        ).fetchone()
        old_container = old_unit["container_id"] if old_unit else None
        restore = _restore_status(conn, old_serial, old_container)
    conn.execute("UPDATE units SET status=? WHERE serial_number=?", (restore, old_serial))

    # Claim new unit
    conn.execute(
        "UPDATE units SET status=? WHERE id=?",
        (f"Allocated ({order_number})", new_unit["id"]),
    )

    if source == "warehouse":
        wh = conn.execute(
            "SELECT id FROM warehouse WHERE order_number=? AND serial_number=?",
            (order_number, old_serial),
        ).fetchone()
        if not wh:
            raise ValueError(f"No warehouse row for {order_number} / {old_serial}")
        conn.execute("""
            UPDATE warehouse SET variant_id=?, serial_number=?, container_id=?
            WHERE id=?
        """, (new_unit["variant_id"], new_unit["serial_number"], new_unit["container_id"], wh["id"]))
    else:
        so = conn.execute(
            "SELECT id FROM sales_orders WHERE order_number=? AND serial_number=?",
            (order_number, old_serial),
        ).fetchone()
        if not so:
            raise ValueError(f"No sales order row for {order_number} / {old_serial}")
        conn.execute("""
            UPDATE sales_orders SET variant_id=?, serial_number=?, container_id=?
            WHERE id=?
        """, (new_unit["variant_id"], new_unit["serial_number"], new_unit["container_id"], so["id"]))

    conn.commit()


def get_warehouse_grouped(conn):
    """Return (order_summaries, unit_rows) for the warehouse page."""
    summaries = conn.execute("""
        SELECT wh.order_number, wh.customer,
               COUNT(*) AS unit_count,
               MIN(wh.date_arrived) AS date_arrived,
               MAX(wh.fulfillment_type) AS fulfillment_type,
               GROUP_CONCAT(v.design_name, ' / ') AS designs,
               CASE
                 WHEN SUM(CASE WHEN wh.status='In Prep' THEN 1 ELSE 0 END) > 0 THEN 'In Prep'
                 WHEN SUM(CASE WHEN wh.status='Pending Review' THEN 1 ELSE 0 END) > 0 THEN 'Pending Review'
                 ELSE MAX(wh.status)
               END AS status
        FROM warehouse wh JOIN variants v ON v.id = wh.variant_id
        GROUP BY wh.order_number
        ORDER BY MIN(wh.id) DESC
    """).fetchall()
    units = conn.execute("""
        SELECT wh.id, wh.order_number, wh.serial_number, wh.container_id, wh.status,
               wh.fulfillment_type, v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku
        FROM warehouse wh JOIN variants v ON v.id = wh.variant_id
        ORDER BY wh.order_number, wh.id
    """).fetchall()
    return summaries, units


def get_sales_grouped(conn):
    """Return (order_summaries, unit_rows) for the sales page — both tables."""
    summaries = conn.execute("""
        SELECT order_number, customer, date_allocated, unit_count, status, designs, source, fulfillment_type
        FROM (
            SELECT so.order_number, so.customer, so.date_allocated,
                   COUNT(*) AS unit_count,
                   MAX(so.status) AS status,
                   GROUP_CONCAT(v.design_name, ' / ') AS designs,
                   'sales' AS source,
                   MAX(so.fulfillment_type) AS fulfillment_type
            FROM sales_orders so JOIN variants v ON v.id = so.variant_id
            GROUP BY so.order_number
            UNION ALL
            SELECT wh.order_number, wh.customer, wh.date_arrived AS date_allocated,
                   COUNT(*) AS unit_count,
                   MAX(wh.status) AS status,
                   GROUP_CONCAT(v.design_name, ' / ') AS designs,
                   'warehouse' AS source,
                   MAX(wh.fulfillment_type) AS fulfillment_type
            FROM warehouse wh JOIN variants v ON v.id = wh.variant_id
            GROUP BY wh.order_number
        )
        ORDER BY order_number
    """).fetchall()
    units = conn.execute("""
        SELECT id, order_number, serial_number, container_id, status,
               design_name, size, finish, swing, glass_type, sku, source, fulfillment_type
        FROM (
            SELECT so.id, so.order_number, so.serial_number, so.container_id, so.status,
                   v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku,
                   'sales' AS source, so.fulfillment_type
            FROM sales_orders so JOIN variants v ON v.id = so.variant_id
            UNION ALL
            SELECT wh.id, wh.order_number, wh.serial_number, wh.container_id, wh.status,
                   v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku,
                   'warehouse' AS source, wh.fulfillment_type
            FROM warehouse wh JOIN variants v ON v.id = wh.variant_id
        )
        ORDER BY order_number, id
    """).fetchall()
    return summaries, units


def get_available_units(conn):
    """Return all In Stock and Pre-Sale units available for allocation, ordered In Stock first."""
    return conn.execute("""
        SELECT u.id, v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku,
               u.serial_number, u.status, u.container_id
        FROM units u
        JOIN variants v ON v.id = u.variant_id
        WHERE u.status = 'In Stock' OR u.status LIKE 'Pre-Sale%'
        ORDER BY
            CASE WHEN u.status = 'In Stock' THEN 0 ELSE 1 END,
            v.design_name, v.size, v.finish
    """).fetchall()


def get_all_containers_in_sales(conn):
    """Return distinct container_ids that have allocated sales orders OR pre-sale units."""
    rows = conn.execute("""
        SELECT DISTINCT container_id FROM sales_orders
        WHERE status = 'Allocated' AND container_id IS NOT NULL
        UNION
        SELECT DISTINCT container_id FROM units
        WHERE status LIKE 'Pre-Sale%' AND container_id IS NOT NULL
        ORDER BY container_id
    """).fetchall()
    return [r["container_id"] for r in rows]


def get_variants_with_negative_variance(conn):
    """Return summary rows where (in_stock + in_prod + pre_sale) < optimal_count."""
    summary = get_units_summary(conn)
    return [r for r in summary if (r["in_stock"] + r["in_prod"] + r["pre_sale"]) < r["optimal_count"]]


def set_optimal_count(conn, variant_id, optimal):
    """Update optimal_count for a variant."""
    conn.execute("UPDATE variants SET optimal_count=? WHERE id=?", (optimal, variant_id))
    conn.commit()


def get_variants_for_review(conn):
    """
    Return all variants with current stock levels, total units ever sold,
    and the earliest sale date — used for quarterly review and forecasting.
    """
    return conn.execute("""
        SELECT
            v.id,
            v.design_name,
            v.size,
            v.finish,
            v.swing,
            v.glass_type,
            v.sku,
            v.optimal_count,
            SUM(CASE WHEN u.status = 'In Stock'      THEN 1 ELSE 0 END) AS in_stock,
            SUM(CASE WHEN u.status = 'In Production' THEN 1 ELSE 0 END) AS in_prod,
            SUM(CASE WHEN u.status LIKE 'Pre-Sale%'  THEN 1 ELSE 0 END) AS pre_sale,
            (
                SELECT COUNT(*) FROM sales_orders so WHERE so.variant_id = v.id
            ) + (
                SELECT COUNT(*) FROM warehouse wh WHERE wh.variant_id = v.id
            ) AS total_sold,
            (
                SELECT MIN(date_allocated) FROM sales_orders so WHERE so.variant_id = v.id
            ) AS first_sale_date
        FROM variants v
        LEFT JOIN units u ON u.variant_id = v.id
        GROUP BY v.id
        ORDER BY v.design_name, v.size, v.finish
    """).fetchall()


def get_all_units(conn):
    """
    Return all units joined with variants, ordered by design_name, size, finish, swing, glass_type.
    """
    return conn.execute("""
        SELECT
            u.id AS unit_id,
            v.id AS variant_id,
            v.design_name,
            v.size,
            v.finish,
            v.swing,
            v.glass_type,
            v.sku,
            v.optimal_count,
            u.serial_number,
            u.status,
            u.container_id,
            u.date_received
        FROM units u
        JOIN variants v ON v.id = u.variant_id
        WHERE u.status NOT LIKE 'Allocated%'
        ORDER BY v.design_name, v.size, v.finish, v.swing, v.glass_type,
                 CASE WHEN u.serial_number IS NULL THEN 1 ELSE 0 END,
                 u.serial_number
    """).fetchall()


def delete_production_units(conn):
    """Remove all In Production units (for cleanup/reset)."""
    conn.execute("DELETE FROM units WHERE status='In Production'")
    conn.commit()


# ── Change requests ────────────────────────────────────────────────────────────

def create_change_request(conn, order_number, customer, order_type, scenario_id,
                          request_detail, notes, created_by):
    from datetime import datetime
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur = conn.execute(
        """INSERT INTO change_requests
           (order_number, customer, order_type, scenario_id, request_detail, notes,
            status, created_by, created_at, updated_at)
           VALUES (?,?,?,?,?,?,'Open',?,?,?)""",
        (order_number, customer, order_type, scenario_id, request_detail, notes,
         created_by, now, now),
    )
    conn.commit()
    return cur.lastrowid


def get_all_change_requests(conn):
    return conn.execute(
        "SELECT * FROM change_requests ORDER BY id DESC"
    ).fetchall()


def get_change_request(conn, cr_id):
    return conn.execute(
        "SELECT * FROM change_requests WHERE id=?", (cr_id,)
    ).fetchone()


def get_warehouse_change_requests(conn):
    """Return open/in-progress change requests for orders currently in the warehouse."""
    return conn.execute("""
        SELECT cr.id, cr.order_number, cr.customer, cr.scenario_id, cr.request_detail,
               cr.status, cr.created_at,
               MAX(wh.status) AS warehouse_status
        FROM change_requests cr
        JOIN warehouse wh ON wh.order_number = cr.order_number
        WHERE cr.status != 'Resolved'
        GROUP BY cr.id
        ORDER BY cr.id DESC
    """).fetchall()


def update_change_request(conn, cr_id, status, resolution, notes):
    from datetime import datetime
    conn.execute(
        """UPDATE change_requests
           SET status=?, resolution=?, notes=?, updated_at=?
           WHERE id=?""",
        (status, resolution, notes, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), cr_id),
    )
    conn.commit()
