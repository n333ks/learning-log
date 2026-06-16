import csv
from datetime import date

# ── CONFIG ───────────────────────────────────────────
MANIFEST_FILE = "container_manifest.csv"
LOG_FILE = "receiving_log.txt"

# Columns that together identify a unique product variant
VARIANT_KEYS = ["Design Name", "Size", "Finish", "Glass Type", "Top Shape", "Hardware", "Swing"]

# ── LOAD MANIFEST ────────────────────────────────────
def load_manifest(filename):
    items = []
    with open(filename, newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            items.append(row)
    return items

# ── BUILD VARIANT LABEL ──────────────────────────────
def variant_label(row):
    parts = [row[k] for k in VARIANT_KEYS if row.get(k) and row[k] != "N/A" and row[k] != "None"]
    return " | ".join(parts)

# ── GENERATE REPORT ──────────────────────────────────
def generate_report(items):
    today = date.today().strftime("%B %d, %Y")
    lines = []

    lines.append("=" * 60)
    lines.append("   CONTAINER RECEIVING REPORT")
    lines.append(f"   Date: {today}")
    lines.append(f"   Total Line Items: {len(items)}")
    total_units = sum(int(row["Quantity"]) for row in items)
    lines.append(f"   Total Units: {total_units}")
    lines.append("=" * 60)
    lines.append("")

    lines.append("ACTION REQUIRED")
    lines.append("-" * 60)
    lines.append("For each item below:")
    lines.append("  1. Remove quantity from PRESALE in Shopify")
    lines.append("  2. Draft the presale product listing")
    lines.append("  3. Add quantity to IN-STOCK version in Shopify")
    lines.append("")

    # Group by design name
    designs = {}
    for row in items:
        name = row["Design Name"]
        if name not in designs:
            designs[name] = []
        designs[name].append(row)

    for design, rows in designs.items():
        lines.append(f"PRODUCT: {design}")
        lines.append("-" * 40)
        design_total = 0
        for row in rows:
            qty = int(row["Quantity"])
            design_total += qty
            label = variant_label(row)
            add_on = row.get("Add-Ons", "None")
            lines.append(f"  ▸ {label}")
            lines.append(f"    Qty to move: {qty} units")
            if add_on and add_on != "None":
                lines.append(f"    Add-on: {add_on}  (+${row.get('Add-On Price', '0')} per unit)")
            lines.append("")
        lines.append(f"  Subtotal for {design}: {design_total} units")
        lines.append("")

    lines.append("=" * 60)
    lines.append(f"TOTAL UNITS RECEIVED: {total_units}")
    lines.append("=" * 60)
    lines.append("")
    lines.append("Received by: _______________________")
    lines.append("Verified by: _______________________")
    lines.append("Notes: ____________________________")

    return "\n".join(lines)

# ── MAIN ─────────────────────────────────────────────
def main():
    print("Loading manifest...")
    items = load_manifest(MANIFEST_FILE)
    print(f"Found {len(items)} line items.\n")

    report = generate_report(items)

    # Print to terminal
    print(report)

    # Save to file
    with open(LOG_FILE, "w") as f:
        f.write(report)

    print(f"\nReport saved to {LOG_FILE}")

if __name__ == "__main__":
    main()
