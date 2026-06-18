# Steel Door Inventory Management

A Python-based inventory management system for a steel door and window 
import/distribution operation. Built around a central Excel workbook 
(`inventory_master.xlsx`) that tracks every physical unit from presale 
through container arrival to in-stock.

## The Problem

Managing inventory across presale, in-transit, and in-stock states for 
a product catalogue with dozens of variants (design × size × finish × 
swing × glass type) was entirely manual. Container arrivals required 
cross-referencing manifests against presale counts, updating quantities 
by hand, and reconciling serial numbers — a process that took hours and 
was prone to error.

## The Solution

A suite of Python scripts that automate inventory tracking around a 
structured Excel workbook:

- **Container arrivals** are processed from a manifest CSV in seconds
- **Serial numbers** are generated and tracked per unit
- **Live COUNTIFS formulas** keep QTY columns accurate automatically
- **Conditional formatting** signals inventory health at a glance

---

## Workbook Structure — `inventory_master.xlsx`

Every row in the workbook represents one physical unit. Columns A–L:

| Column | Field | Notes |
|--------|-------|-------|
| A | Design Name | e.g. Valencia, Sevilla |
| B | Size | e.g. 36" × 80" |
| C | Finish | e.g. Matte Black |
| D | Swing | e.g. Right Inswing |
| E | Glass Type | e.g. Clear, Frosted |
| F | In-Stock QTY | COUNTIFS formula (summary row only) |
| G | In-Production QTY | COUNTIFS formula (summary row only) |
| H | Pre-Sale QTY | COUNTIFS formula (summary row only) |
| I | Optimal Count | Target stock level (summary row only) |
| J | Variance | =F+G+H−I (summary row only) |
| K | Serial Number | Per-unit identifier |
| L | Status | In Stock / Pre-Sale / In Production / Allocated - [order#] |

### Row layout

- **Row 1**: Column headers
- **Row 2**: Blank spacer
- **Data rows**: one per physical unit, grouped by variant

Within each variant group the first row is the **summary row** — it 
carries the COUNTIFS formulas, Optimal Count, and Variance. All 
subsequent rows in the group (**detail rows**) contain only Serial 
Number and Status; the variant label cells (A–E) retain their values 
but display in white font so the repeated text is visually hidden.

Spacing between groups:
- 1 blank row between variants within the same Design Name
- 2 blank rows between Design Name blocks

### Conditional formatting

| Status value | Fill color |
|---|---|
| In Stock | Green |
| Pre-Sale | Yellow |
| In Production | Red |
| Allocated - [any] | Blue |
| Variance < 0 | Red fill + red font |

---

## Scripts

### Day-to-day operations

**`inventory_update.py`** — Process a container arrival.  
Reads `container_manifest.csv`, decrements the matching container's 
Pre-Sale QTY, and moves units to In Stock. Writes serial numbers to 
new detail rows if provided.

```bash
python3 inventory_update.py --container CNT-001
python3 inventory_update.py --container CNT-002 --manifest my_manifest.csv
```

The manifest CSV requires these columns:
`Design Name, Size, Finish, Swing, Glass Type, Quantity, Serial Numbers`
(Serial Numbers is semicolon-separated, optional.)

---

**`fix_inventory_formatting.py`** — Reapply borders and conditional formatting.  
Run this after manually adding or removing rows in Excel to redraw 
design-block borders and ensure CF ranges are correct.

```bash
python3 fix_inventory_formatting.py
```

---

### Setup / rebuild

**`rebuild_inventory.py`** — Full workbook rebuild.  
Reads the current file, preserves all In Stock serials, and writes a 
fresh `inventory_master.xlsx` with the correct flat per-unit structure, 
COUNTIFS formulas, CF rules, and borders.

```bash
python3 rebuild_inventory.py
```

---

### One-time migration scripts

These were run in sequence to build the current workbook and should not 
need to be run again under normal operation. They are kept for reference.

| Script | Purpose |
|--------|---------|
| `create_inventory_master.py` | Generated the original template with sample product data |
| `seed_initial_inventory.py` | Seeded In Stock serial numbers for products with existing stock |
| `add_row_spacing.py` | Inserted blank rows between variant groups and Design Name blocks |
| `fix_summary_rows.py` | Corrected COUNTIFS formula placement (summary rows only) and fixed row reference offsets introduced by row insertions |
| `fix_detail_rows.py` | Set white font on repeated variant labels in detail rows; cleared Optimal Count from detail rows |
| `apply_status_fills.py` | Early attempt at static status fills — superseded by CF rules in `fix_inventory_formatting.py` |

---

## Tech Stack

- Python 3
- openpyxl (Excel workbook read/write, conditional formatting, borders)
- Built with AI-assisted development
