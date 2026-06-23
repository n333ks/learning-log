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
- Units within each variant group are ordered by serial number ascending (FIFO — earliest serial first)
- Excel workbook (`inventory_master.xlsx`) stays in sync after every DB mutation

### Sales
- All orders grouped by order number with expand/collapse to see individual units
- **Unit Swap**: swap a physical unit on any order — old unit returns to In Stock (warehouse orders) or Pre-Sale/In Stock (pre-arrival orders) based on whether its container has arrived
- **Cancel Order**: restores unit to correct status
- **Change Request**: log a customer-initiated change request directly from any order row (see Change Requests below)
- Fulfillment type badge (🏠 Pickup / 🚚 Delivery) shown on every order

### Warehouse
Active change requests for orders currently in the warehouse are shown in a yellow alert banner at the top of the page — warehouse staff see them before anything else.

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
- "Mark Ready" button appears when all 14 items are checked → sets unit to Pending Review

### Order Summary Page
Per-order summary at `/warehouse/order/<order_number>`:
- All units in the order with individual progress bars and notes
- All photos from all units in one grid, labeled by serial number
- Links to each unit's individual prep sheet

### Change Requests
Tracks customer-initiated change requests through a structured procedure. Accessible at `/change-requests` (admin only).

**How it works:**
- Click "Change Request" on any order in the Sales tab
- The app auto-detects the order's current location in the pipeline and resolves the applicable scenario without asking unnecessary questions
- For orders still in house (In Prep, Pending Review, or still in sales_orders), scenario **1A** is auto-assigned
- For orders marked Ready for Pickup or Ready for Delivery, a single follow-up question determines the scenario

**Scenarios (Stock / Presale Orders):**
| Scenario | Situation | Outcome |
|---|---|---|
| 1A | Still in the warehouse | Change can be made |
| 1B | Shipped via LTL | Reversal possible — fees apply |
| 1C | Shipped via non-LTL | No reversal — 25% restocking fee |
| 1D | Picked up or locally delivered | 25% restocking fee |

Each change request detail page shows:
- CS steps, warehouse actions, and owner verification checklist for the applicable scenario
- The current units on the order with a **Swap Unit** picker to execute the inventory change in the same view
- Status tracking (Open → In Progress → Resolved) with resolution notes
- Non-negotiables reminder

Active (unresolved) change requests for warehouse orders surface automatically on the Warehouse page.

### Activity Log (Admin)
Append-only audit trail at `/admin/activity`:
- Logs every checklist action, photo upload/delete, notes save, Mark Ready, cancel, unit swap, change request creation/update, and sign-in
- Filterable by user

### Users (Admin)
- Two roles: `admin` (full access) and `warehouse` (warehouse tab only, read-only on change requests)
- Create users, reset passwords, delete users at `/admin/users`

### Containers
- Upload container manifests (CSV) to parse serials and update inventory
- Receive a container to move Allocated units → Warehouse and Pre-Sale units → In Stock

### Purchase Orders
- Auto-generated from variants with negative variance (stock below optimal count)
- Seeds In Production units for each line item
- Downloadable as CSV with unit cost and retail price

### Quarterly Review
- Forecasts suggested optimal counts per variant based on sales velocity
- Confidence rating (High / Medium / Low / No data) based on days of sales history
- Editable optimal counts saved directly to the DB

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
| `change_requests` | Customer change request log with scenario, procedure steps, and resolution |
| `users` | Auth — admin or warehouse role |
| `activity_log` | Audit trail |

---

## Serial Number Format

```
YY-MMDD-####
```

Year ordered, date (MMDD), sequential number resetting each year. Example: `26-0618-0023` = ordered in 2026, June 18, sequential #23. Units are displayed and fulfilled in serial number order (FIFO).

---

## SKU Format

```
DESIGN-WxH-FINISH-SWING-GLASS
```

Example: `VAL-36X80-MB-RI-CLR` = Valencia Single Door, 36×80, Matte Black, Right Inswing, Clear.

Design codes: ALT · CAD · COR · GRN · MAL · MAR · RON · SEG · SEV · TOL · VAL

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
