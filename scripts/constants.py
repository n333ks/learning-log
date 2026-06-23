import os

# Paths — resolved relative to this file so scripts work from any directory
_SCRIPTS_DIR     = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR        = os.path.dirname(_SCRIPTS_DIR)          # inventory_management/
INVENTORY_FILE   = os.path.join(_BASE_DIR, "inventory_master.xlsx")
MANIFEST_DIR     = os.path.join(_BASE_DIR, "manifests")
MANIFEST_PATTERN = os.path.join(MANIFEST_DIR, "container_manifest_CNT-*.csv")
DB_PATH          = os.path.join(_BASE_DIR, "data", "inventory.db")

# Column indices (1-based)
DATA_START  = 3
VARIANT_COLS = range(1, 6)   # A–E

C_SKU      = 6
C_INSTOCK  = 7
C_INPROD   = 8
C_PRESALE  = 9
C_OPTIMAL  = 10
C_VARIANCE = 11
C_SERIAL   = 12
C_STATUS   = 13

TOTAL_COLS = 13   # A through M

SKU_MAP = {
    ("Altea Double Door",      '32" x 80"', "Satin White",       "Right Inswing",  "Clear"):    "ALT-32X80-SW-RI-CLR",
    ("Altea Double Door",      '36" x 80"', "Matte Black",       "Left Inswing",   "Rain"):     "ALT-36X80-MB-LI-RN",
    ("Altea Double Door",      '36" x 96"', "Oil-Rubbed Bronze", "Right Outswing", "Sidelite"): "ALT-36X96-ORB-RO-SDL",
    ("Altea Double Door",      '72" x 80"', "Matte Black",       "Left Inswing",   "Clear"):    "ALT-72X80-MB-LI-CLR",
    ("Cadiz Pivot Door",       '42" x 84"', "Matte Black",       "Left Inswing",   "Clear"):    "CAD-42X84-MB-LI-CLR",
    ("Cadiz Pivot Door",       '48" x 84"', "Brushed Nickel",    "Left Inswing",   "Frosted"):  "CAD-48X84-BN-LI-FRS",
    ("Cordova Sliding Window", '32" x 80"', "Matte Black",       "Right Inswing",  "Frosted"):  "COR-32X80-MB-RI-FRS",
    ("Cordova Sliding Window", '36" x 80"', "Oil-Rubbed Bronze", "Left Inswing",   "Clear"):    "COR-36X80-ORB-LI-CLR",
    ("Cordova Sliding Window", '36" x 96"', "Brushed Nickel",    "Right Outswing", "Rain"):     "COR-36X96-BN-RO-RN",
    ("Cordova Sliding Window", '48" x 36"', "Matte Black",       "Left Inswing",   "Clear"):    "COR-48X36-MB-LI-CLR",
    ("Granada Bi-fold Door",   '36" x 80"', "Matte Black",       "Right Inswing",  "Clear"):    "GRN-36X80-MB-RI-CLR",
    ("Granada Bi-fold Door",   '36" x 96"', "Satin White",       "Left Inswing",   "Sidelite"): "GRN-36X96-SW-LI-SDL",
    ("Granada Bi-fold Door",   '96" x 80"', "Matte Black",       "Left Inswing",   "Clear"):    "GRN-96X80-MB-LI-CLR",
    ("Malaga Single Door",     '32" x 80"', "Brushed Nickel",    "Left Inswing",   "None"):     "MAL-32X80-BN-LI-NON",
    ("Malaga Single Door",     '36" x 80"', "Gunmetal",          "Left Inswing",   "None"):     "MAL-36X80-GM-LI-NON",
    ("Ronda Louvered Door",    '32" x 80"', "Satin White",       "Right Inswing",  "None"):     "RON-32X80-SW-RI-NON",
    ("Ronda Louvered Door",    '36" x 80"', "Matte Black",       "Left Inswing",   "None"):     "RON-36X80-MB-LI-NON",
    ("Sevilla French Door",    '32" x 80"', "Brushed Nickel",    "Right Inswing",  "Rain"):     "SEV-32X80-BN-RI-RN",
    ("Sevilla French Door",    '36" x 80"', "Satin White",       "Left Inswing",   "None"):     "SEV-36X80-SW-LI-NON",
    ("Sevilla French Door",    '72" x 80"', "Matte White",       "Left Inswing",   "Frosted"):  "SEV-72X80-MW-LI-FRS",
    ("Toledo Dutch Door",      '36" x 80"', "Gunmetal",          "Right Inswing",  "Rain"):     "TOL-36X80-GM-RI-RN",
    ("Toledo Dutch Door",      '36" x 80"', "Matte Black",       "Left Inswing",   "Clear"):    "TOL-36X80-MB-LI-CLR",
    ("Valencia Single Door",   '32" x 80"', "Matte Black",       "Left Inswing",   "Clear"):    "VAL-32X80-MB-LI-CLR",
    ("Valencia Single Door",   '36" x 80"', "Matte Black",       "Right Inswing",  "Clear"):    "VAL-36X80-MB-RI-CLR",
    ("Valencia Single Door",   '36" x 96"', "Oil-Rubbed Bronze", "Right Outswing", "Frosted"):  "VAL-36X96-ORB-RO-FRS",
}

_FINISH_ABBREV = {
    "matte black":       "MB",
    "brushed nickel":    "BN",
    "oil-rubbed bronze": "ORB",
    "satin white":       "SW",
    "gunmetal":          "GM",
    "matte white":       "MW",
    "brushed gold":      "BG",
    "antique brass":     "AB",
    "chrome":            "CHR",
    "flat black":        "FB",
}

_SWING_ABBREV = {
    "right inswing":  "RI",
    "left inswing":   "LI",
    "right outswing": "RO",
    "left outswing":  "LO",
}

_GLASS_ABBREV = {
    "clear":    "CLR",
    "frosted":  "FRS",
    "rain":     "RN",
    "sidelite": "SDL",
    "none":     "NON",
    "obscure":  "OBS",
    "bronze":   "BRZ",
}


def _design_prefix(design_name):
    """First 3 uppercase letters of the first word of the design name."""
    first_word = str(design_name).strip().split()[0]
    letters = [c for c in first_word.upper() if c.isalpha()]
    return ''.join(letters[:3])


def _size_code(size):
    """'36" x 80"' → '36X80'"""
    return str(size).replace('"', '').replace(' x ', 'X').replace(' X ', 'X').strip()


def _abbrev(value, lookup):
    """Look up abbreviation; fall back to uppercase initials of each word."""
    key = str(value).strip().lower()
    if key in lookup:
        return lookup[key]
    return ''.join(w[0].upper() for w in key.split() if w)


def generate_sku(design, size, finish, swing, glass):
    """Dynamically build a SKU from variant attributes."""
    prefix = _design_prefix(design)
    sz     = _size_code(size)
    fin    = _abbrev(finish, _FINISH_ABBREV)
    sw     = _abbrev(swing,  _SWING_ABBREV)
    gl     = _abbrev(glass,  _GLASS_ABBREV)
    return f"{prefix}-{sz}-{fin}-{sw}-{gl}"


# ── Cost calculator ───────────────────────────────────────────────────────────
_COST_RATES = {
    'pivot':   40,
    'bi-fold': 50,
    'sliding': 55,
}
_DEFAULT_RATE = 30   # all swing types: single, double, french, louvered, dutch, barn

def get_rate(design_name):
    """Return cost per sq ft based on door type."""
    name = str(design_name).lower()
    for keyword, rate in _COST_RATES.items():
        if keyword in name:
            return rate
    return _DEFAULT_RATE


def parse_size(size):
    """'32" x 80"' → (32, 80). Returns (0, 0) if unparseable."""
    try:
        parts = str(size).replace('"', '').split('x')
        if len(parts) != 2:
            parts = str(size).replace('"', '').split('X')
        return int(parts[0].strip()), int(parts[1].strip())
    except Exception:
        return 0, 0


def calculate_cost(design_name, size):
    """
    Return (unit_cost, retail_price) based on sq footage and door type.
    Retail is 4x cost. Both rounded to 2 decimal places.
    """
    w, h        = parse_size(size)
    sq_footage  = (w * h) / 144
    rate        = get_rate(design_name)
    unit_cost   = round(sq_footage * rate, 2)
    retail      = round(unit_cost * 4, 2)
    return unit_cost, retail


def get_sku(variant_key):
    """Return SKU for a (Design, Size, Finish, Swing, Glass) tuple.
    Uses the curated SKU_MAP first; falls back to dynamic generation."""
    key = tuple(str(v).strip() for v in variant_key)
    if key in SKU_MAP:
        return SKU_MAP[key]
    return generate_sku(*key)

# Sales tab columns (1-based)
S_ORDER     = 1
S_CUSTOMER  = 2
S_DESIGN    = 3
S_SIZE      = 4
S_FINISH    = 5
S_SWING     = 6
S_GLASS     = 7
S_SKU       = 8
S_SERIAL    = 9
S_CONTAINER = 10
S_DATE      = 11
S_STATUS    = 12

# Warehouse tab columns (same layout, Date Arrived instead of Date Allocated)
W_ORDER     = 1
W_CUSTOMER  = 2
W_DESIGN    = 3
W_SIZE      = 4
W_FINISH    = 5
W_SWING     = 6
W_GLASS     = 7
W_SKU       = 8
W_SERIAL    = 9
W_CONTAINER = 10
W_DATE      = 11
W_STATUS    = 12

SALES_TAB     = "Sales"
WAREHOUSE_TAB = "Warehouse"
