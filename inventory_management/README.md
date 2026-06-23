# Inventory Management System

Flask-based inventory and warehouse management for an iron doors business. Tracks every unit from production order through container arrival, customer allocation, warehouse prep, and final fulfillment.

---

## Running the App

```bash
cd inventory_management
python3 app.py
```

Opens at `http://localhost:5001`. Default admin login: `admin` / `admin`.

Dependencies: `flask`, `werkzeug`, `openpyxl`

---

## Unit Lifecycle

```
In Production → Pre-Sale (CNT-XXX) → Allocated → Warehouse (In Prep) → Pending Review → Ready for Pickup / Ready for Delivery
```

---

## Features

### Inventory
- Tracks every physical unit with status, serial number, container, and variant details
- Allocated units are hidden from the inventory page — only In Stock, Pre-Sale, and In Production are shown
- Excel workbook (`inventory_master.xlsx`) stays in sync after every DB mutation

### Sales
- All orders grouped by order number with expand/collapse to see individual units
- **Change Order**: swap a unit on any order (pickup or delivery) — old unit returns to Pre-Sale or In Stock based on whether its container has arrived
- **Cancel Order**: restores unit to correct status; no In Production row created
- Fulfillment type badge (🏠 Pickup / 🚚 Delivery) shown on every order

### Warehouse
Four sections updated in real time:
1. **In Prep** — warehouse team works on units, accesses individual prep sheets
2. **Pending Review** — all units marked ready; office staff reviews before release
3. **Ready for Pickup** — approved pickup orders
4. **Ready for Delivery** — approved delivery orders

Office staff can **Accept** (moves to Ready) or **Push Back** (returns to In Prep).

### Warehouse Prep Sheets
Each unit gets its own prep sheet at `/warehouse/prep/<id>`:
- 14-item QC checklist with instant AJAX save
- Photo uploads with lightbox viewer
- Notes with auto-save on blur
- Progress bar (checks completed / total)
- "Mark Ready for Pickup" button appears when all 14 items are checked → sets unit to Pending Review

### Order Summary Page
Per-order summary at `/warehouse/order/<order_number>`:
- All units in the order with individual progress bars and notes
- All photos from all units in one grid, labeled by serial number
- Links to each unit's individual prep sheet

### Activity Log (Admin)
Append-only audit trail at `/admin/activity`:
- Logs every checklist check/uncheck, photo upload/delete, notes save, Mark Ready, Cancel, Change Order, and sign-in
- Each entry links directly to the relevant prep sheet
- Filterable by user

### Users (Admin)
- Two roles: `admin` (full access) and `warehouse` (warehouse tab only, no cancel/change)
- Create users, reset passwords, delete users at `/admin/users`

### Containers
- Upload container manifests (CSV) to parse serials and update inventory
- Receive a container to move Allocated units → Warehouse and Pre-Sale units → In Stock

---

## Database Tables

| Table | Purpose |
|---|---|
| `units` | Every physical unit |
| `variants` | Unique design/size/finish/swing/glass combos |
| `sales_orders` | Pre-arrival customer allocations |
| `warehouse` | Units that have arrived and are being prepped |
| `warehouse_checklist` | 14-item QC checklist per unit |
| `warehouse_photos` | Photos per unit |
| `users` | Auth — admin or warehouse role |
| `activity_log` | Audit trail |

---

## Legacy Excel Scripts

Scripts in `scripts/` predate the Flask app and operate directly on `inventory_master.xlsx`. Still useful for bulk operations:

```bash
python3 scripts/parse_manifest.py       # parse container manifest → update Excel
python3 scripts/sort_inventory.py       # re-sort design blocks alphabetically
python3 scripts/refresh_counts.py       # recalculate QTY counts and variance
python3 scripts/apply_status_fills.py   # reapply colors + column widths
```

`scripts/constants.py` is the single source of truth for file paths, column indices, and the SKU map.

---

## SKU Format

```
DESIGN-WxH-FINISH-SWING-GLASS
```

Example: `VAL-36X80-MB-RI-CLR` = Valencia Single Door, 36×80, Matte Black, Right Inswing, Clear.

Design codes: ALT · CAD · COR · GRN · MAL · RON · SEV · TOL · VAL
