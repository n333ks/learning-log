# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the App

```bash
python3 app.py          # starts Flask on http://localhost:5001
```

Run from the repo root. The `scripts/` directory is added to `sys.path` at startup. Default admin credentials: `admin` / `admin`.

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
- **Excel layer** (`scripts/` + `inventory_master.xlsx`) — legacy scripts for bulk ops, kept in sync
- **Web layer** (`app.py` + `data/inventory.db`) — primary day-to-day system

`app.py` calls `export_excel(conn)` after every DB mutation so the Excel file stays current.

### Request flow
All DB access goes through `scripts/db.py`. `app.py` imports everything from there — no raw SQL in routes. `scripts/constants.py` is the single source of truth for file paths, column indices, and the SKU map.

### Unit lifecycle
```
In Production → Pre-Sale (CNT-XXX) → Allocated (ORD-XXX) → Warehouse (In Prep) → Pending Review → Ready for Pickup / Ready for Delivery
```

- Units with `status LIKE 'Allocated%'` are filtered out of the inventory page.
- `_restore_status(conn, serial, container_id)` determines the correct status when a unit is released from an order. For warehouse-source swaps, always restore to `In Stock`. For sales-source swaps, checks `date_received` on the unit.
- `move_to_warehouse()` copies a `sales_orders` row to `warehouse` and deletes the source row.
- Units are ordered by serial number ascending within each variant group (FIFO). Serial format: `YY-MMDD-####`.

### Schema tables

| Table | Purpose |
|---|---|
| `units` | Every physical unit; statuses: `In Stock`, `Pre-Sale (CNT-XXX)`, `In Production`, `Allocated (ORD-XXX)` |
| `variants` | Unique (design, size, finish, swing, glass) combos with SKU and optimal count |
| `sales_orders` | Active pre-arrival allocations |
| `warehouse` | Units physically on-site being prepped |
| `warehouse_checklist` | 14-item QC checklist per warehouse unit (keyed on `warehouse_id`); column is `completed` (not `checked`) |
| `warehouse_photos` | Photos per warehouse unit |
| `change_requests` | Customer change request log — scenario, procedure steps, status, resolution |
| `order_details` | Customer contact info (phone, email, address) per order number |
| `order_photos` | Photos attached to a pipeline order (separate from warehouse prep photos) |
| `containers` | Container ETAs keyed by `container_id` |
| `users` | Roles: `admin` or `warehouse` |
| `activity_log` | Append-only audit trail |

`init_db()` creates all tables and runs pending `ALTER TABLE` migrations wrapped in try/except. New columns must be added to both the `SCHEMA` string and the migrations list.

### Sales Pipeline (`/pipeline`)
Admin-only kanban board across 7 stages: `presale → allocated → in_prep → pending_review → ready_pickup → ready_delivery → fulfilled`.

- Stage is derived in `get_pipeline_orders(conn)` — presale if container hasn't arrived, allocated if in `sales_orders` with arrived container, in_prep/pending_review/ready if in `warehouse`.
- `arrived` field: `CASE WHEN so.container_id IS NULL THEN 1 WHEN EXISTS (SELECT 1 FROM units u2 WHERE u2.container_id=so.container_id AND u2.status='In Stock') THEN 1 ELSE 0 END`
- Pipeline cards link with `?stage=` param so the detail page can build the stage sidebar.
- `get_pipeline_order_detail(conn, order_number)` returns the full dict including `units` (with `checklist`, `checklist_done`, `checklist_total`, `prep_photos`, `container_eta`), `change_orders` (grouped by timestamp), `activity`, `customer`, `days_in_stage`, `first_wh_id`, and `timeline`.
- Checklist progress for warehouse orders: `checklist_total = unit_count × 14`.
- Stage sidebar in `pipeline_detail.html` uses `position: fixed; left: 240px; width: 216px` — sits immediately right of the 240px main nav. Content area uses `margin-left: 216px`.

### Change Orders (`scripts/change_orders.py`)
All 7 scenario procedures (1A, 1B, 1C, 1D, B1C, B1A, B1B) live in `SCENARIOS` dict. `determine_scenario()` resolves the scenario from order type + follow-up answers. The `/change-orders/new` route auto-detects the scenario — no manual input needed for most cases. The detail page shows CS steps, warehouse actions, owner verification, and an inline unit swap picker that posts to `/sales/change-order`. Unit swap displays include swing and glass type everywhere (change order detail, warehouse alerts, pipeline detail).

### Auth and roles
`login_required` — any logged-in user. `admin_required` — admin only. Warehouse users redirect to `/warehouse` on login and cannot see other tabs. Use `generate_password_hash(password, method='pbkdf2:sha256')` — `scrypt` is not available on this system's OpenSSL.

### Templates
All templates extend `base.html`. Grouped order views use a Jinja `unit_map` dict built via `namespace` to map `order_number → [units]`. Expand/collapse via `toggleOrder()` in inline `<script>` blocks. AJAX calls (checklist, notes) post to `/warehouse/prep/<wh_id>/<action>` and return `{"ok": true}`.

### Key conventions
- Warehouse prep photo uploads: `static/uploads/<wh_id>/<timestamped_filename>`.
- Pipeline order photo uploads: `static/uploads/orders/<order_number>/<timestamped_filename>`.
- `fulfillment_type` (`pickup` / `delivery`) lives on both `sales_orders` and `warehouse`. Currently randomly assigned at startup; Shopify will set it in the future.
- Warehouse change requests surface automatically on the warehouse page — `get_warehouse_change_requests()` joins `change_requests` with `warehouse` on `order_number` and returns only unresolved ones.

## Scripts

Active scripts imported by `app.py` live in `scripts/`:
- `db.py` — all DB access
- `constants.py` — file paths, column indices, SKU map
- `export_excel.py` — syncs SQLite → `inventory_master.xlsx`
- `parse_manifest.py` — parses container manifest CSVs
- `change_orders.py` — change request procedure data and scenario logic

One-off migration scripts that predate the Flask app are in `scripts/legacy/` and are not imported anywhere.
