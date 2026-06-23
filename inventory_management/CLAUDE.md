# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the App

```bash
cd inventory_management
python3 app.py          # starts Flask on http://localhost:5001
```

Flask must be run from the `inventory_management/` directory. The `scripts/` directory is added to `sys.path` at startup. Default admin credentials: `admin` / `admin`.

To test routes without a browser:
```python
import sys; sys.path.insert(0, 'scripts')
from db import init_db; init_db()
from app import app
app.config['TESTING'] = True
with app.test_client() as c:
    with c.session_transaction() as sess:
        sess['logged_in'] = True; sess['username'] = 'admin'
        sess['role'] = 'admin'; sess['full_name'] = 'Admin'
        sess['user'] = {'username': 'admin', 'role': 'admin', 'full_name': 'Admin'}
    r = c.get('/some-route')
```

## Architecture

### Two parallel systems
- **Excel layer** (`scripts/` + `inventory_master.xlsx`) — legacy scripts for bulk ops, still kept in sync
- **Web layer** (`app.py` + `inventory.db`) — primary day-to-day system

`app.py` calls `export_excel(conn)` after every DB mutation so the Excel file stays current.

### Request flow
All DB access goes through `scripts/db.py`. `app.py` imports everything from there — no raw SQL in routes. `scripts/constants.py` is the single source of truth for file paths, column indices, and the SKU map.

### Unit lifecycle
```
In Production → Pre-Sale (CNT-XXX) → Allocated (ORD-XXX) → Warehouse (In Prep) → Pending Review → Ready for Pickup / Ready for Delivery
```

- Units with `status LIKE 'Allocated%'` are filtered out of the inventory page.
- `_restore_status(conn, serial, container_id)` determines the correct status when a unit is released from an order back to inventory. For warehouse-source swaps, always restore to `In Stock` (container has physically arrived). For sales-source swaps, use `_restore_status` which checks `date_received` on the unit.
- `move_to_warehouse()` copies a `sales_orders` row to `warehouse` and deletes the source row.
- Units are ordered by serial number ascending within each variant group (FIFO). Serial format: `YY-MMDD-####`.

### Schema tables

| Table | Purpose |
|---|---|
| `units` | Every physical unit; statuses: `In Stock`, `Pre-Sale (CNT-XXX)`, `In Production`, `Allocated (ORD-XXX)` |
| `variants` | Unique (design, size, finish, swing, glass) combos with SKU and optimal count |
| `sales_orders` | Active pre-arrival allocations |
| `warehouse` | Units physically on-site being prepped |
| `warehouse_checklist` | 14-item QC checklist per warehouse unit (keyed on `warehouse_id`) |
| `warehouse_photos` | Photos per warehouse unit |
| `change_requests` | Customer change request log — scenario, procedure steps, status, resolution |
| `users` | Roles: `admin` or `warehouse` |
| `activity_log` | Append-only audit trail |

`init_db()` creates all tables and runs pending `ALTER TABLE` migrations wrapped in try/except. New columns must be added to both the `SCHEMA` string and the migrations list.

### Change requests (`scripts/change_orders.py`)
All 7 scenario procedures (1A, 1B, 1C, 1D, B1C, B1A, B1B) live in `SCENARIOS` dict. `determine_scenario()` resolves the scenario from order type + follow-up answers. The `/change-requests/new` route auto-detects the scenario by querying `sales_orders` and `warehouse` — no manual input needed for most cases. The detail page shows CS steps, warehouse actions, owner verification, and an inline unit swap picker that posts to the existing `/sales/change-order` endpoint.

### Auth and roles
`login_required` — any logged-in user. `admin_required` — admin only. Warehouse users redirect to `/warehouse` on login and cannot see other tabs. Use `generate_password_hash(password, method='pbkdf2:sha256')` — `scrypt` is not available on this system's OpenSSL.

### Templates
All templates extend `base.html`. Grouped order views use a Jinja `unit_map` dict built via `namespace` to map `order_number → [units]`. Expand/collapse via `toggleOrder()` in inline `<script>` blocks. AJAX calls (checklist, notes) post to `/warehouse/prep/<wh_id>/<action>` and return `{"ok": true}`.

### Key conventions
- Photo uploads stored at `static/uploads/<wh_id>/<timestamped_filename>`.
- `fulfillment_type` (`pickup` / `delivery`) lives on both `sales_orders` and `warehouse`. Currently randomly assigned at startup; Shopify will set it in the future.
- Warehouse change requests surface automatically on the warehouse page — `get_warehouse_change_requests()` joins `change_requests` with `warehouse` on `order_number` and returns only unresolved ones.

## Legacy Excel Scripts

```bash
python3 scripts/parse_manifest.py        # parse a container manifest CSV → updates Excel
python3 scripts/sort_inventory.py        # re-sort design blocks alphabetically
python3 scripts/refresh_counts.py        # recalculate QTY counts and variance
python3 scripts/apply_status_fills.py    # reapply status colors + column widths
```

Dependency: `openpyxl`. These are largely superseded by the Flask app but kept for bulk Excel operations.
