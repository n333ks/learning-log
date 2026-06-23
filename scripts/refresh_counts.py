#!/usr/bin/env python3
"""
Thin compatibility wrapper.

Counts are always derived from inventory.db — this module exists so existing
callers (sort_inventory, etc.) can still import refresh_counts without errors.

When called standalone it triggers an Excel export.
"""

import os
import sys

_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SCRIPTS_DIR)


def refresh_counts(conn=None):
    """
    No-op compatibility shim.
    Counts are always computed from the DB by get_units_summary().
    Pass conn to optionally trigger an export_excel call.
    """
    # Kept for backwards compat — callers may pass an openpyxl worksheet;
    # if conn is an sqlite3 Connection we can trigger a re-export.
    import sqlite3
    if conn is not None and isinstance(conn, sqlite3.Connection):
        from export_excel import export_excel
        export_excel(conn)


def main():
    from db import get_conn
    from export_excel import export_excel
    conn = get_conn()
    print("Refreshing counts (regenerating Excel from DB) ...")
    export_excel(conn)
    conn.close()
    print("Done.")


if __name__ == "__main__":
    main()
