#!/usr/bin/env python3
"""
Database layer for the inventory management system.
All scripts import from here instead of using openpyxl directly.
"""

import sqlite3
import os
from datetime import datetime

from constants import DB_PATH

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
    fulfillment_type TEXT DEFAULT 'pickup',
    awaiting_presale INTEGER DEFAULT 0
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

CREATE TABLE IF NOT EXISTS change_request_units (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    cr_id         INTEGER NOT NULL REFERENCES change_requests(id),
    action        TEXT NOT NULL,
    serial_number TEXT,
    design_name   TEXT,
    size          TEXT,
    finish        TEXT,
    swing         TEXT,
    glass_type    TEXT,
    sku           TEXT
);

CREATE TABLE IF NOT EXISTS change_requests (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    order_number     TEXT NOT NULL,
    customer         TEXT NOT NULL,
    order_type       TEXT NOT NULL,
    scenario_id      TEXT NOT NULL,
    request_detail   TEXT,
    notes            TEXT,
    status           TEXT DEFAULT 'Open',
    resolution       TEXT,
    created_by       TEXT,
    created_at       TEXT,
    updated_at       TEXT,
    warehouse_ack    INTEGER DEFAULT 0,
    warehouse_ack_by TEXT,
    warehouse_ack_at TEXT
);

CREATE TABLE IF NOT EXISTS order_details (
    order_number  TEXT PRIMARY KEY,
    phone         TEXT,
    email         TEXT,
    address       TEXT,
    city          TEXT,
    state         TEXT,
    zip           TEXT,
    notes         TEXT
);

CREATE TABLE IF NOT EXISTS order_photos (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    order_number  TEXT NOT NULL,
    filename      TEXT NOT NULL,
    caption       TEXT,
    uploaded_by   TEXT,
    uploaded_at   TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS containers (
    container_id  TEXT PRIMARY KEY,
    eta           TEXT,
    notes         TEXT
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
        "ALTER TABLE sales_orders ADD COLUMN awaiting_presale INTEGER DEFAULT 0",
        "ALTER TABLE change_requests ADD COLUMN warehouse_ack INTEGER DEFAULT 0",
        "ALTER TABLE change_requests ADD COLUMN warehouse_ack_by TEXT",
        "ALTER TABLE change_requests ADD COLUMN warehouse_ack_at TEXT",
        "ALTER TABLE change_request_units ADD COLUMN created_at TEXT",
        """CREATE TABLE IF NOT EXISTS order_details (
            order_number TEXT PRIMARY KEY, phone TEXT, email TEXT,
            address TEXT, city TEXT, state TEXT, zip TEXT, notes TEXT)""",
        """CREATE TABLE IF NOT EXISTS order_photos (
            id INTEGER PRIMARY KEY AUTOINCREMENT, order_number TEXT NOT NULL,
            filename TEXT NOT NULL, caption TEXT, uploaded_by TEXT,
            uploaded_at TEXT DEFAULT (datetime('now')))""",
        """CREATE TABLE IF NOT EXISTS containers (
            container_id TEXT PRIMARY KEY, eta TEXT, notes TEXT)""",
    ]:
        try:
            conn.execute(sql)
        except Exception:
            pass
    conn.commit()
    conn.close()


# ── Activity log ──────────────────────────────────────────────────────────────

def log_activity(conn, username, full_name, action, detail=None, order_number=None, serial_number=None, warehouse_id=None):
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
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
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
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
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


# ── Variants ──────────────────────────────────────────────────────────────────

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


# ── Units ─────────────────────────────────────────────────────────────────────

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


# ── Sales orders ──────────────────────────────────────────────────────────────

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
    return f"Pre-Sale - {container_id}" if container_id else "In Stock"


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
        # Warehouse units may not have a row in units table — upsert
        existing = conn.execute(
            "SELECT id FROM units WHERE serial_number=?", (old_serial,)
        ).fetchone()
        if existing:
            conn.execute("UPDATE units SET status=? WHERE serial_number=?", (restore, old_serial))
        else:
            # Re-insert from warehouse row info
            wh_row = conn.execute(
                "SELECT variant_id, container_id FROM warehouse WHERE serial_number=? AND order_number=?",
                (old_serial, order_number),
            ).fetchone()
            if wh_row:
                conn.execute(
                    "INSERT INTO units (variant_id, serial_number, status, container_id) VALUES (?, ?, ?, ?)",
                    (wh_row["variant_id"], old_serial, restore, wh_row["container_id"]),
                )
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
            "SELECT id, customer, fulfillment_type FROM warehouse WHERE order_number=? AND serial_number=?",
            (order_number, old_serial),
        ).fetchone()
        if not wh:
            raise ValueError(f"No warehouse row for {order_number} / {old_serial}")

        if new_unit["status"].startswith("Pre-Sale"):
            # Move this warehouse unit back to sales_orders (awaiting arrival)
            conn.execute("""
                INSERT INTO sales_orders
                    (order_number, customer, variant_id, serial_number, container_id,
                     date_allocated, status, fulfillment_type, awaiting_presale)
                VALUES (?, ?, ?, ?, ?, date('now'), 'Allocated', ?, 1)
            """, (order_number, wh["customer"], new_unit["variant_id"],
                  new_unit["serial_number"], new_unit["container_id"],
                  wh["fulfillment_type"]))
            # Remove checklist/photos for the departing warehouse row
            conn.execute("DELETE FROM warehouse_checklist WHERE warehouse_id=?", (wh["id"],))
            conn.execute("DELETE FROM warehouse_photos WHERE warehouse_id=?", (wh["id"],))
            conn.execute("DELETE FROM warehouse WHERE id=?", (wh["id"],))
        else:
            conn.execute("""
                UPDATE warehouse SET variant_id=?, serial_number=?, container_id=?
                WHERE id=?
            """, (new_unit["variant_id"], new_unit["serial_number"], new_unit["container_id"], wh["id"]))
    else:
        so = conn.execute(
            "SELECT * FROM sales_orders WHERE order_number=? AND serial_number=?",
            (order_number, old_serial),
        ).fetchone()
        if not so:
            raise ValueError(f"No sales order row for {order_number} / {old_serial}")
        new_is_presale = new_unit["status"].startswith("Pre-Sale")
        if new_is_presale:
            # Stay in sales_orders, flag as awaiting pre-sale arrival
            conn.execute("""
                UPDATE sales_orders SET variant_id=?, serial_number=?, container_id=?,
                    awaiting_presale=1
                WHERE id=?
            """, (new_unit["variant_id"], new_unit["serial_number"],
                  new_unit["container_id"], so["id"]))
        else:
            # New unit is In Stock — update this row first
            conn.execute("""
                UPDATE sales_orders SET variant_id=?, serial_number=?, container_id=?,
                    awaiting_presale=0
                WHERE id=?
            """, (new_unit["variant_id"], new_unit["serial_number"],
                  new_unit["container_id"], so["id"]))
            # Only move to warehouse if every unit's container has physically arrived
            # (i.e. each container_id has at least one In Stock unit)
            unready = conn.execute("""
                SELECT COUNT(*) FROM sales_orders so
                WHERE so.order_number=?
                  AND NOT EXISTS (
                      SELECT 1 FROM units u
                      WHERE u.container_id = so.container_id
                        AND u.status = 'In Stock'
                  )
            """, (order_number,)).fetchone()[0]
            if unready == 0:
                from datetime import date
                today = date.today().isoformat()
                all_rows = conn.execute(
                    "SELECT * FROM sales_orders WHERE order_number=?", (order_number,)
                ).fetchall()
                for row in all_rows:
                    ft = row["fulfillment_type"] or "pickup"
                    conn.execute("""
                        INSERT INTO warehouse
                            (order_number, customer, variant_id, serial_number, container_id,
                             date_arrived, status, fulfillment_type)
                        VALUES (?, ?, ?, ?, ?, ?, 'In Prep', ?)
                    """, (order_number, row["customer"], row["variant_id"],
                          row["serial_number"], row["container_id"], today, ft))
                conn.execute("DELETE FROM sales_orders WHERE order_number=?", (order_number,))
                # Auto-acknowledge any open COs — warehouse is seeing this order fresh
                from datetime import datetime as _dt
                conn.execute("""
                    UPDATE change_requests SET warehouse_ack=1, warehouse_ack_by='system',
                        warehouse_ack_at=?
                    WHERE order_number=? AND (warehouse_ack IS NULL OR warehouse_ack=0)
                """, (_dt.now().strftime("%Y-%m-%d %H:%M:%S"), order_number))

    conn.commit()


# ── Inventory queries ─────────────────────────────────────────────────────────

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
               design_name, size, finish, swing, glass_type, sku, source, fulfillment_type, arrived
        FROM (
            SELECT so.id, so.order_number, so.serial_number, so.container_id, so.status,
                   v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku,
                   'sales' AS source, so.fulfillment_type,
                   CASE
                       WHEN so.container_id IS NULL THEN 1
                       WHEN EXISTS (
                           SELECT 1 FROM units u2
                           WHERE u2.container_id = so.container_id AND u2.status = 'In Stock'
                       ) THEN 1
                       ELSE 0
                   END AS arrived
            FROM sales_orders so JOIN variants v ON v.id = so.variant_id
            UNION ALL
            SELECT wh.id, wh.order_number, wh.serial_number, wh.container_id, wh.status,
                   v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku,
                   'warehouse' AS source, wh.fulfillment_type, 1 AS arrived
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
            v.design_name, v.size, v.finish, v.swing, v.glass_type,
            CASE WHEN u.status = 'In Stock' THEN 0 ELSE 1 END,
            u.serial_number
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
                 CASE u.status
                     WHEN 'In Stock'      THEN 0
                     WHEN 'In Production' THEN 1
                     ELSE 2
                 END,
                 CASE WHEN u.serial_number IS NULL THEN 1 ELSE 0 END,
                 u.serial_number
    """).fetchall()


def delete_production_units(conn):
    """Remove all In Production units (for cleanup/reset)."""
    conn.execute("DELETE FROM units WHERE status='In Production'")
    conn.commit()


# ── Order mutations ───────────────────────────────────────────────────────────

def remove_unit_from_order(conn, order_number, serial, source):
    """Remove a single unit from an order and restore it to inventory."""
    if source == 'warehouse':
        wh = conn.execute(
            "SELECT id FROM warehouse WHERE order_number=? AND serial_number=?",
            (order_number, serial)
        ).fetchone()
        if wh:
            conn.execute("DELETE FROM warehouse_checklist WHERE warehouse_id=?", (wh['id'],))
            conn.execute("DELETE FROM warehouse_photos WHERE warehouse_id=?", (wh['id'],))
            conn.execute("DELETE FROM warehouse WHERE id=?", (wh['id'],))
        conn.execute("UPDATE units SET status='In Stock' WHERE serial_number=?", (serial,))
    else:
        so = conn.execute(
            "SELECT container_id FROM sales_orders WHERE order_number=? AND serial_number=?",
            (order_number, serial)
        ).fetchone()
        container = so['container_id'] if so else None
        restore = _restore_status(conn, serial, container)
        conn.execute("UPDATE units SET status=? WHERE serial_number=?", (restore, serial))
        conn.execute(
            "DELETE FROM sales_orders WHERE order_number=? AND serial_number=?",
            (order_number, serial)
        )
    conn.commit()


def add_unit_to_order(conn, order_number, customer, new_unit_id, source):
    """Add a new unit to an existing order. Returns a dict of the added unit with variant info."""
    unit = conn.execute("""
        SELECT u.id, u.variant_id, u.serial_number, u.status, u.container_id,
               v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku
        FROM units u JOIN variants v ON v.id = u.variant_id
        WHERE u.id=?
    """, (new_unit_id,)).fetchone()
    if not unit:
        raise ValueError(f"Unit {new_unit_id} not found")

    is_presale = unit['status'].startswith('Pre-Sale')
    conn.execute(
        "UPDATE units SET status=? WHERE id=?",
        (f"Allocated ({order_number})", unit['id'])
    )

    if source == 'warehouse' and not is_presale:
        wh_row = conn.execute(
            "SELECT fulfillment_type FROM warehouse WHERE order_number=? LIMIT 1", (order_number,)
        ).fetchone()
        ft = wh_row['fulfillment_type'] if wh_row else 'pickup'
        conn.execute("""
            INSERT INTO warehouse
                (order_number, customer, variant_id, serial_number, container_id,
                 date_arrived, status, fulfillment_type)
            VALUES (?, ?, ?, ?, ?, date('now'), 'In Prep', ?)
        """, (order_number, customer, unit['variant_id'], unit['serial_number'],
              unit['container_id'], ft))
    else:
        ft_row = conn.execute(
            "SELECT fulfillment_type FROM sales_orders WHERE order_number=? LIMIT 1", (order_number,)
        ).fetchone()
        if not ft_row:
            ft_row = conn.execute(
                "SELECT fulfillment_type FROM warehouse WHERE order_number=? LIMIT 1", (order_number,)
            ).fetchone()
        ft = ft_row['fulfillment_type'] if ft_row else 'pickup'
        awaiting = 1 if is_presale else 0
        conn.execute("""
            INSERT INTO sales_orders
                (order_number, customer, variant_id, serial_number, container_id,
                 date_allocated, status, fulfillment_type, awaiting_presale)
            VALUES (?, ?, ?, ?, ?, date('now'), 'Allocated', ?, ?)
        """, (order_number, customer, unit['variant_id'], unit['serial_number'],
              unit['container_id'], ft, awaiting))

    conn.commit()
    return dict(unit)


def log_cr_unit_change(conn, cr_id, action, unit_info, swap_ts=None):
    """Log a unit added/removed event on a change request."""
    ts = swap_ts or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute("""
        INSERT INTO change_request_units
            (cr_id, action, serial_number, design_name, size, finish, swing, glass_type, sku, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (cr_id, action,
          unit_info.get('serial_number'), unit_info.get('design_name'),
          unit_info.get('size'), unit_info.get('finish'), unit_info.get('swing'),
          unit_info.get('glass_type'), unit_info.get('sku'), ts))
    conn.commit()


def get_cr_unit_changes(conn, cr_ids):
    """Return a dict {cr_id: [change_rows]} for the given list of CR IDs."""
    if not cr_ids:
        return {}
    placeholders = ','.join('?' * len(cr_ids))
    rows = conn.execute(
        f"SELECT * FROM change_request_units WHERE cr_id IN ({placeholders}) ORDER BY id",
        cr_ids
    ).fetchall()
    result = {cid: [] for cid in cr_ids}
    for r in rows:
        result[r['cr_id']].append(r)
    return result


def get_presale_pending_orders(conn):
    """Return sales_order rows where a CO swap demoted the order back from warehouse to pre-sale.
    Only shown when there is an open (unresolved, unacknowledged) change request for the order."""
    return conn.execute("""
        SELECT so.order_number, so.customer, so.serial_number, so.container_id,
               so.fulfillment_type,
               v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku
        FROM sales_orders so
        JOIN variants v ON v.id = so.variant_id
        WHERE so.awaiting_presale = 1
          AND EXISTS (
              SELECT 1 FROM change_requests cr
              WHERE cr.order_number = so.order_number
                AND cr.status != 'Resolved'
                AND (cr.warehouse_ack IS NULL OR cr.warehouse_ack = 0)
          )
        ORDER BY so.order_number, so.id
    """).fetchall()


# ── Change requests ────────────────────────────────────────────────────────────

def create_change_request(conn, order_number, customer, order_type, scenario_id,
                          request_detail, notes, created_by):
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
    """Return unresolved, unacknowledged change requests for orders currently in the warehouse."""
    return conn.execute("""
        SELECT cr.id, cr.order_number, cr.customer, cr.scenario_id, cr.request_detail,
               cr.status, cr.created_at,
               MAX(wh.status) AS warehouse_status
        FROM change_requests cr
        JOIN warehouse wh ON wh.order_number = cr.order_number
        WHERE cr.status != 'Resolved' AND (cr.warehouse_ack IS NULL OR cr.warehouse_ack = 0)
        GROUP BY cr.id
        ORDER BY cr.id DESC
    """).fetchall()


def ack_warehouse_change_request(conn, cr_id, username):
    conn.execute(
        "UPDATE change_requests SET warehouse_ack=1, warehouse_ack_by=?, warehouse_ack_at=? WHERE id=?",
        (username, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), cr_id),
    )
    conn.commit()


def get_units_on_order(conn, order_number):
    """Return units currently on an order from sales_orders or warehouse, joined with variant info."""
    rows = conn.execute("""
        SELECT so.serial_number, so.container_id, 'sales' AS source,
               v.id AS variant_id, v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku,
               CASE
                   WHEN so.awaiting_presale=1 THEN 'Awaiting Pre-Sale'
                   WHEN (SELECT COUNT(*) FROM units u2
                         WHERE u2.container_id = so.container_id
                           AND u2.status = 'In Stock') > 0
                        THEN 'Allocated · In Stock'
                   WHEN so.container_id IS NOT NULL THEN 'Allocated · Pre-Sale'
                   ELSE so.status
               END AS status
        FROM sales_orders so JOIN variants v ON v.id = so.variant_id
        WHERE so.order_number = ?
        UNION ALL
        SELECT wh.serial_number, wh.container_id, 'warehouse' AS source,
               v.id AS variant_id, v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku,
               wh.status AS status
        FROM warehouse wh JOIN variants v ON v.id = wh.variant_id
        WHERE wh.order_number = ?
        ORDER BY serial_number
    """, (order_number, order_number)).fetchall()
    return rows


def update_change_request(conn, cr_id, status, resolution, notes):
    conn.execute(
        """UPDATE change_requests
           SET status=?, resolution=?, notes=?, updated_at=?
           WHERE id=?""",
        (status, resolution, notes, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), cr_id),
    )
    conn.commit()


# ── Sales Pipeline ─────────────────────────────────────────────────────────────

def get_pipeline_orders(conn):
    """Return a list of dicts, one per order, with pipeline_stage assigned."""
    sales_rows = conn.execute("""
        SELECT so.order_number,
               MAX(so.customer) AS customer,
               MAX(so.fulfillment_type) AS fulfillment_type,
               COUNT(*) AS unit_count,
               GROUP_CONCAT(DISTINCT v.design_name) AS designs,
               MIN(so.date_allocated) AS date_allocated,
               GROUP_CONCAT(DISTINCT so.container_id) AS container_ids
        FROM sales_orders so
        JOIN variants v ON v.id = so.variant_id
        GROUP BY so.order_number
    """).fetchall()

    orders = []
    for row in sales_rows:
        # Pre-sale if any unit has a container that hasn't arrived yet
        presale = conn.execute("""
            SELECT 1 FROM sales_orders so
            WHERE so.order_number=?
              AND so.container_id IS NOT NULL
              AND NOT EXISTS (
                  SELECT 1 FROM units u WHERE u.container_id=so.container_id AND u.status='In Stock'
              )
            LIMIT 1
        """, (row['order_number'],)).fetchone()
        stage = 'presale' if presale else 'allocated'
        orders.append({
            'order_number': row['order_number'],
            'customer': row['customer'],
            'fulfillment_type': row['fulfillment_type'],
            'unit_count': row['unit_count'],
            'designs': row['designs'],
            'date_allocated': row['date_allocated'],
            'pipeline_stage': stage,
        })

    wh_rows = conn.execute("""
        SELECT wh.order_number,
               MAX(wh.customer) AS customer,
               MAX(wh.fulfillment_type) AS fulfillment_type,
               COUNT(*) AS unit_count,
               GROUP_CONCAT(DISTINCT v.design_name) AS designs,
               MIN(wh.date_arrived) AS date_allocated,
               GROUP_CONCAT(wh.status) AS all_statuses
        FROM warehouse wh
        JOIN variants v ON v.id = wh.variant_id
        GROUP BY wh.order_number
    """).fetchall()

    # Checklist progress per warehouse order
    checklist_progress = {}
    cl_rows = conn.execute("""
        SELECT wh.order_number,
               COUNT(cl.id)                                    AS total,
               SUM(CASE WHEN cl.completed=1 THEN 1 ELSE 0 END) AS done
        FROM warehouse wh
        LEFT JOIN warehouse_checklist cl ON cl.warehouse_id = wh.id
        GROUP BY wh.order_number
    """).fetchall()
    for r in cl_rows:
        unit_count_row = next((w for w in wh_rows if w['order_number'] == r['order_number']), None)
        units = unit_count_row['unit_count'] if unit_count_row else 1
        # total possible = units × 14 items; cl rows only exist when checklist initialised
        checklist_progress[r['order_number']] = {
            'done': r['done'] or 0,
            'total': units * 14,
        }

    for row in wh_rows:
        statuses = [s.strip() for s in (row['all_statuses'] or '').split(',')]
        if all(s in ('Ready for Pickup', 'Ready for Delivery') for s in statuses):
            ft = (row['fulfillment_type'] or '').lower()
            stage = 'ready_delivery' if ft == 'delivery' else 'ready_pickup'
        elif all(s == 'Pending Review' for s in statuses):
            stage = 'pending_review'
        else:
            stage = 'in_prep'
        cl = checklist_progress.get(row['order_number'], {'done': 0, 'total': row['unit_count'] * 14})
        orders.append({
            'order_number': row['order_number'],
            'customer': row['customer'],
            'fulfillment_type': row['fulfillment_type'],
            'unit_count': row['unit_count'],
            'designs': row['designs'],
            'date_allocated': row['date_allocated'],
            'pipeline_stage': stage,
            'checklist_done': cl['done'],
            'checklist_total': cl['total'],
        })

    return orders


def get_pipeline_order_detail(conn, order_number):
    """Return full detail for a pipeline order card."""
    order_info = conn.execute(
        "SELECT * FROM order_details WHERE order_number=?", (order_number,)
    ).fetchone()

    # Units — include warehouse.id and notes so we can fetch checklists and photos
    unit_rows = conn.execute("""
        SELECT serial_number, design_name, size, finish, swing, glass_type, sku,
               container_id, status, source, arrived, warehouse_id, date_label, notes
        FROM (
            SELECT so.serial_number, v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku,
                   so.container_id, so.status, 'sales' AS source,
                   CASE
                       WHEN so.container_id IS NULL THEN 1
                       WHEN EXISTS (SELECT 1 FROM units u2 WHERE u2.container_id=so.container_id AND u2.status='In Stock') THEN 1
                       ELSE 0
                   END AS arrived,
                   NULL AS warehouse_id,
                   so.date_allocated AS date_label,
                   NULL AS notes
            FROM sales_orders so JOIN variants v ON v.id = so.variant_id
            WHERE so.order_number = ?
            UNION ALL
            SELECT wh.serial_number, v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku,
                   wh.container_id, wh.status, 'warehouse' AS source, 1 AS arrived,
                   wh.id AS warehouse_id,
                   wh.date_arrived AS date_label,
                   wh.notes AS notes
            FROM warehouse wh JOIN variants v ON v.id = wh.variant_id
            WHERE wh.order_number = ?
        )
        ORDER BY serial_number
    """, (order_number, order_number)).fetchall()

    # Checklists and warehouse photos keyed by warehouse_id
    wh_ids = [r['warehouse_id'] for r in unit_rows if r['warehouse_id'] is not None]
    checklist_by_wh = {}
    wh_photos_by_wh = {}
    if wh_ids:
        ph = ','.join('?' * len(wh_ids))
        cl_rows = conn.execute(
            f"SELECT * FROM warehouse_checklist WHERE warehouse_id IN ({ph})", wh_ids
        ).fetchall()
        for r in cl_rows:
            checklist_by_wh.setdefault(r['warehouse_id'], {})[r['item_key']] = bool(r['completed'])
        wp_rows = conn.execute(
            f"SELECT * FROM warehouse_photos WHERE warehouse_id IN ({ph}) ORDER BY id", wh_ids
        ).fetchall()
        for r in wp_rows:
            wh_photos_by_wh.setdefault(r['warehouse_id'], []).append(dict(r))

    units = []
    for r in unit_rows:
        u = dict(r)
        wh_id = u.get('warehouse_id')
        if wh_id is not None:
            cl = checklist_by_wh.get(wh_id, {})
            done = sum(1 for v in cl.values() if v)
            u['checklist'] = [
                {'key': key, 'label': label, 'checked': cl.get(key, False)}
                for key, label in CHECKLIST_ITEMS
            ]
            u['checklist_done'] = done
            u['checklist_total'] = len(CHECKLIST_ITEMS)
            u['prep_photos'] = wh_photos_by_wh.get(wh_id, [])
        else:
            u['checklist'] = []
            u['checklist_done'] = 0
            u['checklist_total'] = 0
            u['prep_photos'] = []
        units.append(u)

    # Order-level photos
    photos = conn.execute(
        "SELECT * FROM order_photos WHERE order_number=? ORDER BY id", (order_number,)
    ).fetchall()

    # Change orders with unit change log
    change_orders = conn.execute(
        "SELECT * FROM change_requests WHERE order_number=? ORDER BY id", (order_number,)
    ).fetchall()
    cr_ids = [cr['id'] for cr in change_orders]
    unit_changes_by_cr = get_cr_unit_changes(conn, cr_ids)
    co_list = []
    for cr in change_orders:
        raw_changes = unit_changes_by_cr.get(cr['id'], [])
        groups = {}
        for ch in raw_changes:
            ts = ch['created_at'] or ''
            groups.setdefault(ts, []).append(dict(ch))
        co_list.append({
            'id': cr['id'],
            'status': cr['status'],
            'created_at': cr['created_at'],
            'created_by': cr['created_by'],
            'swap_groups': [
                {'timestamp': ts, 'changes': changes}
                for ts, changes in sorted(groups.items())
            ],
        })

    # Activity log
    activity = conn.execute(
        """SELECT full_name, username, action, detail, serial_number, created_at
           FROM activity_log WHERE order_number=? ORDER BY id""",
        (order_number,)
    ).fetchall()

    # Derive summary fields for the banner
    customer = next((u['design_name'] for u in units), None)  # placeholder
    so_meta = conn.execute(
        "SELECT customer, fulfillment_type, MIN(date_allocated) AS date_allocated FROM sales_orders WHERE order_number=? GROUP BY order_number",
        (order_number,)
    ).fetchone()
    wh_meta = conn.execute(
        "SELECT customer, fulfillment_type, MIN(date_arrived) AS date_arrived FROM warehouse WHERE order_number=? GROUP BY order_number",
        (order_number,)
    ).fetchone()
    customer = (so_meta and so_meta['customer']) or (wh_meta and wh_meta['customer']) or ''
    fulfillment_type = (so_meta and so_meta['fulfillment_type']) or (wh_meta and wh_meta['fulfillment_type']) or ''
    date_allocated = so_meta['date_allocated'] if so_meta else None
    date_warehouse = wh_meta['date_arrived'] if wh_meta else None

    # Pipeline stage
    in_sales = bool(so_meta)
    in_wh = bool(wh_meta)
    if in_sales:
        presale = conn.execute("""
            SELECT 1 FROM sales_orders so WHERE so.order_number=?
              AND so.container_id IS NOT NULL
              AND NOT EXISTS (SELECT 1 FROM units u WHERE u.container_id=so.container_id AND u.status='In Stock')
            LIMIT 1
        """, (order_number,)).fetchone()
        stage = 'presale' if presale else 'allocated'
    else:
        statuses = [u['status'] for u in units if u['source'] == 'warehouse']
        if all(s in ('Ready for Pickup', 'Ready for Delivery') for s in statuses):
            stage = 'ready_delivery' if (fulfillment_type or '').lower() == 'delivery' else 'ready_pickup'
        elif all(s == 'Pending Review' for s in statuses):
            stage = 'pending_review'
        else:
            stage = 'in_prep'

    # Timeline: stages with dates where known
    timeline = [
        {'key': 'allocated',      'label': 'Allocated',    'date': date_allocated},
        {'key': 'in_prep',        'label': 'In Prep',      'date': date_warehouse},
        {'key': 'pending_review', 'label': 'Pending Review','date': None},
        {'key': 'ready',          'label': 'Ready',         'date': None},
    ]
    stage_order_keys = ['presale', 'allocated', 'in_prep', 'pending_review', 'ready_pickup', 'ready_delivery']
    current_idx = stage_order_keys.index(stage) if stage in stage_order_keys else 1
    tl_keys = ['allocated', 'in_prep', 'pending_review', 'ready']
    tl_past = {'allocated': 1, 'in_prep': 2, 'pending_review': 3, 'ready_pickup': 4, 'ready_delivery': 4}
    current_tl_level = tl_past.get(stage, 1)
    for item in timeline:
        level = tl_keys.index(item['key']) + 1
        item['done'] = level < current_tl_level
        item['current'] = level == current_tl_level

    # Days in current stage
    from datetime import date as _date
    today = _date.today()
    stage_start_str = date_warehouse if stage in ('in_prep', 'pending_review', 'ready_pickup', 'ready_delivery') else date_allocated
    days_in_stage = None
    if stage_start_str:
        try:
            stage_start = _date.fromisoformat(stage_start_str)
            days_in_stage = (today - stage_start).days
        except ValueError:
            pass

    # Container ETAs keyed by container_id
    container_ids = list({u['container_id'] for u in units if u['container_id']})
    container_etas = {}
    if container_ids:
        ph = ','.join('?' * len(container_ids))
        for r in conn.execute(f"SELECT container_id, eta FROM containers WHERE container_id IN ({ph})", container_ids).fetchall():
            container_etas[r['container_id']] = r['eta']
    # Attach ETA to each unit
    for u in units:
        u['container_eta'] = container_etas.get(u['container_id']) if u['container_id'] else None

    # First warehouse ID for the prep page link (#6)
    first_wh = conn.execute(
        "SELECT MIN(id) AS id FROM warehouse WHERE order_number=?", (order_number,)
    ).fetchone()
    first_wh_id = first_wh['id'] if first_wh else None

    return {
        'order_info': dict(order_info) if order_info else {},
        'units': units,
        'photos': [dict(p) for p in photos],
        'change_orders': co_list,
        'activity': [dict(a) for a in activity],
        'customer': customer,
        'fulfillment_type': fulfillment_type,
        'date_allocated': date_allocated,
        'date_warehouse': date_warehouse,
        'stage': stage,
        'timeline': timeline,
        'unit_count': len(units),
        'days_in_stage': days_in_stage,
        'first_wh_id': first_wh_id,
    }


# ── Order details & photos ────────────────────────────────────────────────────

def upsert_order_details(conn, order_number, phone, email, address, city, state, zip_, notes):
    conn.execute(
        """INSERT OR REPLACE INTO order_details
           (order_number, phone, email, address, city, state, zip, notes)
           VALUES (?,?,?,?,?,?,?,?)""",
        (order_number, phone, email, address, city, state, zip_, notes),
    )
    conn.commit()


def add_order_photo(conn, order_number, filename, caption, uploaded_by):
    conn.execute(
        """INSERT INTO order_photos (order_number, filename, caption, uploaded_by)
           VALUES (?,?,?,?)""",
        (order_number, filename, caption, uploaded_by),
    )
    conn.commit()


def delete_order_photo(conn, photo_id):
    row = conn.execute(
        "SELECT filename, order_number FROM order_photos WHERE id=?", (photo_id,)
    ).fetchone()
    conn.execute("DELETE FROM order_photos WHERE id=?", (photo_id,))
    conn.commit()
    return (row['filename'], row['order_number']) if row else (None, None)
