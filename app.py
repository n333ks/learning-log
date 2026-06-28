#!/usr/bin/env python3
"""
Flask web application for inventory management.
Run with: python3 app.py
"""

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_from_directory, session, g
from functools import wraps
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename
import sys
import os
import glob
import csv
import datetime
import random

# Add scripts directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'scripts'))

from db import (
    get_conn, get_units_summary, get_all_units, get_variants_with_negative_variance,
    get_or_create_variant, count_production_units, fill_production_unit,
    add_unit, add_sales_order, get_all_containers_in_sales,
    get_sales_by_container, move_to_warehouse, cancel_order,
    set_optimal_count, get_variants_for_review,
    change_order_unit, get_available_units, cancel_warehouse_order,
    get_user, get_all_users, create_user, update_user_password, delete_user, seed_admin,
    ensure_checklist, toggle_checklist_item, save_warehouse_notes,
    add_warehouse_photo, delete_warehouse_photo, mark_warehouse_ready, get_prep_data,
    get_warehouse_grouped, get_sales_grouped, get_order_prep_summary,
    log_activity, get_activity_log,
    approve_warehouse_order, reject_warehouse_order,
    create_change_request, get_all_change_requests, get_change_request, update_change_request,
    get_warehouse_change_requests,
    get_presale_pending_orders,
    ack_warehouse_change_request,
    get_units_on_order,
    remove_unit_from_order,
    add_unit_to_order,
    log_cr_unit_change,
    get_cr_unit_changes,
    CHECKLIST_ITEMS,
    get_pipeline_orders, get_pipeline_order_detail,
    upsert_order_details, add_order_photo, delete_order_photo,
)
from change_orders import SCENARIOS, NON_NEGOTIABLES, FRICTION_COLOR, determine_scenario, get_scenario
from constants import MANIFEST_DIR, MANIFEST_PATTERN, DB_PATH, get_sku, calculate_cost
from parse_manifest import load_manifest, manifest_key, manifest_serials, manifest_status, container_id_from_filename, process_manifest
from export_excel import export_excel

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'inventory-secret-key-2024')
app.jinja_env.globals['calculate_cost'] = calculate_cost

BASE_DIR   = os.path.dirname(__file__)
PO_DIR     = os.path.join(BASE_DIR, 'purchase_orders')
UPLOAD_DIR = os.path.join(BASE_DIR, 'static', 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)
ORDER_UPLOAD_DIR = os.path.join(UPLOAD_DIR, 'orders')
os.makedirs(ORDER_UPLOAD_DIR, exist_ok=True)
ALLOWED_EXTENSIONS = {'jpg', 'jpeg', 'png', 'webp', 'heic'}

# Keep legacy env-var credentials for seeding the admin on first run
_LEGACY_ADMIN_USER = os.environ.get('ADMIN_USER', 'admin')
_LEGACY_ADMIN_HASH = os.environ.get('ADMIN_PASS_HASH',
    'pbkdf2:sha256:1000000$fGzn6HnnQh73TBct$b3f1f60d6b92cf60ba7252f579ebf51daa59de965f31a38537df565cdcd851d3')


@app.context_processor
def inject_user():
    return {
        'current_user': session.get('user', {}),
        'is_admin': session.get('role') == 'admin',
    }


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login', next=request.path))
        return f(*args, **kwargs)
    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login', next=request.path))
        if session.get('role') != 'admin':
            flash('Access restricted to administrators.', 'error')
            return redirect(url_for('warehouse'))
        return f(*args, **kwargs)
    return decorated


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def _session_user():
    """Return (username, full_name) from the current session."""
    username = session.get('username', 'unknown')
    return username, session.get('full_name', username)


def _save_photo(file, dest_dir):
    """Validate, save, and return the timestamped filename. Returns None if invalid."""
    if not file or not allowed_file(file.filename):
        return None
    os.makedirs(dest_dir, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    filename = f"{ts}_{secure_filename(file.filename)}"
    file.save(os.path.join(dest_dir, filename))
    return filename


# ── Helpers ────────────────────────────────────────────────────────────────────

def get_warehouse_rows(conn):
    return conn.execute("""
        SELECT wh.id, wh.order_number, wh.customer, v.design_name, v.size, v.finish,
               v.swing, v.glass_type, v.sku, wh.serial_number, wh.container_id,
               wh.date_arrived, wh.status
        FROM warehouse wh
        JOIN variants v ON v.id = wh.variant_id
        ORDER BY wh.id DESC
    """).fetchall()


def count_by_status(conn, status):
    row = conn.execute(
        "SELECT COUNT(*) AS n FROM units WHERE status=?", (status,)
    ).fetchone()
    return row['n']


def count_units_total(conn):
    row = conn.execute("SELECT COUNT(*) AS n FROM units").fetchone()
    return row['n']


def get_recent_sales(conn, limit=5):
    return conn.execute("""
        SELECT order_number, customer, design_name, date_allocated, status
        FROM (
            SELECT so.order_number, so.customer, v.design_name,
                   so.date_allocated, so.status, so.id AS sort_id
            FROM sales_orders so
            JOIN variants v ON v.id = so.variant_id
            UNION ALL
            SELECT wh.order_number, wh.customer, v.design_name,
                   wh.date_arrived AS date_allocated, wh.status, wh.id AS sort_id
            FROM warehouse wh
            JOIN variants v ON v.id = wh.variant_id
        )
        GROUP BY order_number
        ORDER BY MAX(sort_id) DESC
        LIMIT ?
    """, (limit,)).fetchall()


def get_container_breakdown(conn):
    return conn.execute("""
        SELECT
            container_id,
            COUNT(*) AS total,
            SUM(CASE WHEN status = 'In Stock'        THEN 1 ELSE 0 END) AS in_stock,
            SUM(CASE WHEN status LIKE 'Pre-Sale%'    THEN 1 ELSE 0 END) AS pre_sale,
            SUM(CASE WHEN status = 'In Production'   THEN 1 ELSE 0 END) AS in_prod,
            SUM(CASE WHEN status LIKE 'Allocated%'   THEN 1 ELSE 0 END) AS allocated
        FROM units
        WHERE container_id IS NOT NULL
        GROUP BY container_id
        ORDER BY container_id
    """).fetchall()


def list_manifests():
    files = sorted(glob.glob(MANIFEST_PATTERN))
    result = []
    for f in files:
        fname = os.path.basename(f)
        try:
            rows = load_manifest(f)
            count = sum(len(manifest_serials(r)) for r in rows)
        except Exception:
            count = 0
        result.append({'path': f, 'filename': fname, 'unit_count': count})
    return result


def list_po_files():
    if not os.path.isdir(PO_DIR):
        return []
    files = sorted(glob.glob(os.path.join(PO_DIR, '*.csv')), reverse=True)
    result = []
    for f in files:
        stat = os.stat(f)
        mtime = datetime.datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M')
        result.append({'filename': os.path.basename(f), 'date': mtime})
    return result


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('logged_in'):
        dest = url_for('warehouse') if session.get('role') == 'warehouse' else url_for('dashboard')
        return redirect(dest)
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        conn = get_conn()
        user = get_user(conn, username)
        conn.close()
        if user and check_password_hash(user['password_hash'], password):
            session['logged_in'] = True
            session['username']  = username
            session['role']      = user['role']
            session['full_name'] = user['full_name'] or username
            session['user']      = {'username': username, 'role': user['role'], 'full_name': user['full_name'] or username}
            conn2 = get_conn()
            log_activity(conn2, username, user['full_name'] or username, 'Signed in')
            conn2.close()
            dest = url_for('warehouse') if user['role'] == 'warehouse' else (request.form.get('next') or url_for('dashboard'))
            return redirect(dest)
        flash('Invalid username or password.', 'error')
    return render_template('login.html', next=request.args.get('next', ''))


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/')
@login_required
def dashboard():
    conn = get_conn()
    in_stock  = count_by_status(conn, 'In Stock')
    in_prod   = count_by_status(conn, 'In Production')
    pre_sale  = conn.execute("SELECT COUNT(*) AS n FROM units WHERE status LIKE 'Pre-Sale%'").fetchone()['n']
    negative_variants   = get_variants_with_negative_variance(conn)
    alerts              = len(negative_variants)
    recent_sales        = get_recent_sales(conn)
    container_breakdown = get_container_breakdown(conn)
    # Inventory value: sum retail price of all In Stock units
    stock_units = conn.execute("""
        SELECT v.design_name, v.size FROM units u
        JOIN variants v ON v.id = u.variant_id
        WHERE u.status = 'In Stock'
    """).fetchall()
    inventory_value = sum(calculate_cost(u['design_name'], u['size'])[1] for u in stock_units)
    conn.close()
    return render_template('dashboard.html',
                           in_stock=in_stock, in_prod=in_prod, pre_sale=pre_sale,
                           alerts=alerts, negative_variants=negative_variants,
                           recent_sales=recent_sales,
                           container_breakdown=container_breakdown,
                           inventory_value=inventory_value)


@app.route('/inventory')
@login_required
def inventory():
    conn = get_conn()
    units = get_all_units(conn)
    summary_rows = get_units_summary(conn)
    summary = {r['id']: r for r in summary_rows}
    containers = sorted({u['container_id'] for u in units if u['container_id']})
    # Build cost lookup keyed by variant_id
    costs = {}
    for u in units:
        if u['variant_id'] not in costs:
            unit_cost, retail = calculate_cost(u['design_name'], u['size'])
            margin = round((retail - unit_cost) / retail * 100) if retail else 0
            costs[u['variant_id']] = {
                'unit_cost': unit_cost,
                'retail':    retail,
                'margin':    margin,
            }
    conn.close()
    return render_template('inventory.html', units=units, summary=summary,
                           containers=containers, costs=costs)


@app.route('/sales')
@login_required
def sales():
    conn = get_conn()
    summaries, units = get_sales_grouped(conn)
    available = [dict(u) for u in get_available_units(conn)]
    conn.close()
    return render_template('sales.html', summaries=summaries, units=units, available_units=available)


@app.route('/warehouse')
@login_required
def warehouse():
    conn = get_conn()
    summaries, units = get_warehouse_grouped(conn)
    progress_rows = conn.execute("""
        SELECT wh.order_number,
               SUM(CASE WHEN wc.completed=1 THEN 1 ELSE 0 END) AS done
        FROM warehouse wh
        LEFT JOIN warehouse_checklist wc ON wc.warehouse_id = wh.id
        GROUP BY wh.order_number
    """).fetchall()
    order_progress = {r['order_number']: (r['done'] or 0) for r in progress_rows}
    change_requests   = get_warehouse_change_requests(conn)
    presale_pending   = get_presale_pending_orders(conn)
    cr_unit_changes   = get_cr_unit_changes(conn, [cr['id'] for cr in change_requests])
    conn.close()
    return render_template('warehouse.html', summaries=summaries, units=units,
                           order_progress=order_progress, total_checklist=len(CHECKLIST_ITEMS),
                           change_requests=change_requests, scenarios=SCENARIOS,
                           presale_pending=presale_pending, cr_unit_changes=cr_unit_changes)


@app.route('/warehouse/prep/<int:wh_id>')
@login_required
def warehouse_prep(wh_id):
    conn = get_conn()
    ensure_checklist(conn, wh_id)
    unit, checklist, photos = get_prep_data(conn, wh_id)
    conn.close()
    if not unit:
        flash('Unit not found.', 'error')
        return redirect(url_for('warehouse'))
    total = len(checklist)
    done  = sum(1 for c in checklist if c['completed'])
    pct   = int(done / total * 100) if total else 0
    return render_template('warehouse_prep.html',
                           unit=unit, checklist=checklist, photos=photos,
                           total=total, done=done, pct=pct,
                           checklist_items=CHECKLIST_ITEMS)


@app.route('/warehouse/prep/<int:wh_id>/check', methods=['POST'])
@login_required
def warehouse_prep_check(wh_id):
    item_key  = request.form.get('item_key')
    completed = request.form.get('completed') == '1'
    username, full_name = _session_user()
    conn = get_conn()
    toggle_checklist_item(conn, wh_id, item_key, completed, username)
    unit = conn.execute("SELECT order_number, serial_number FROM warehouse WHERE id=?", (wh_id,)).fetchone()
    label = dict(CHECKLIST_ITEMS).get(item_key, item_key)
    action = 'Checked' if completed else 'Unchecked'
    log_activity(conn, username, full_name, f'{action}: {label}',
                 order_number=unit['order_number'] if unit else None,
                 serial_number=unit['serial_number'] if unit else None,
                 warehouse_id=wh_id)
    conn.close()
    return jsonify({'ok': True})


@app.route('/warehouse/prep/<int:wh_id>/notes', methods=['POST'])
@login_required
def warehouse_prep_notes(wh_id):
    notes    = request.form.get('notes', '')
    username, full_name = _session_user()
    conn = get_conn()
    save_warehouse_notes(conn, wh_id, notes)
    unit = conn.execute("SELECT order_number, serial_number FROM warehouse WHERE id=?", (wh_id,)).fetchone()
    log_activity(conn, username, full_name, 'Updated notes',
                 order_number=unit['order_number'] if unit else None,
                 serial_number=unit['serial_number'] if unit else None,
                 warehouse_id=wh_id)
    conn.close()
    return jsonify({'ok': True})


@app.route('/warehouse/prep/<int:wh_id>/photo', methods=['POST'])
@login_required
def warehouse_prep_photo(wh_id):
    caption  = request.form.get('caption', '')
    filename = _save_photo(request.files.get('photo'), os.path.join(UPLOAD_DIR, str(wh_id)))
    if not filename:
        flash('Please upload a JPG, PNG, or WebP image.', 'error')
        return redirect(url_for('warehouse_prep', wh_id=wh_id))
    username, full_name = _session_user()
    conn = get_conn()
    add_warehouse_photo(conn, wh_id, filename, caption, username)
    unit = conn.execute("SELECT order_number, serial_number FROM warehouse WHERE id=?", (wh_id,)).fetchone()
    log_activity(conn, username, full_name, 'Uploaded photo',
                 detail=caption or filename,
                 order_number=unit['order_number'] if unit else None,
                 serial_number=unit['serial_number'] if unit else None,
                 warehouse_id=wh_id)
    conn.close()
    return redirect(url_for('warehouse_prep', wh_id=wh_id))


@app.route('/warehouse/prep/<int:wh_id>/photo/<int:photo_id>/delete', methods=['POST'])
@admin_required
def warehouse_photo_delete(wh_id, photo_id):
    username, full_name = _session_user()
    conn = get_conn()
    unit = conn.execute("SELECT order_number, serial_number FROM warehouse WHERE id=?", (wh_id,)).fetchone()
    filename, _ = delete_warehouse_photo(conn, photo_id)
    log_activity(conn, username, full_name, 'Deleted photo',
                 order_number=unit['order_number'] if unit else None,
                 serial_number=unit['serial_number'] if unit else None,
                 warehouse_id=wh_id)
    conn.close()
    if filename:
        path = os.path.join(UPLOAD_DIR, str(wh_id), filename)
        if os.path.exists(path):
            os.remove(path)
    return redirect(url_for('warehouse_prep', wh_id=wh_id))


@app.route('/warehouse/prep/<int:wh_id>/ready', methods=['POST'])
@login_required
def warehouse_prep_ready(wh_id):
    username, full_name = _session_user()
    conn = get_conn()
    mark_warehouse_ready(conn, wh_id)
    unit = conn.execute("SELECT order_number, serial_number FROM warehouse WHERE id=?", (wh_id,)).fetchone()
    log_activity(conn, username, full_name, 'Marked Ready for Pickup',
                 order_number=unit['order_number'] if unit else None,
                 serial_number=unit['serial_number'] if unit else None,
                 warehouse_id=wh_id)
    conn.close()
    flash('Unit marked as Ready for Pickup.', 'success')
    return redirect(url_for('warehouse_prep', wh_id=wh_id))


@app.route('/warehouse/order/<order_number>/approve', methods=['POST'])
@login_required
def warehouse_order_approve(order_number):
    username, full_name = _session_user()
    conn = get_conn()
    approve_warehouse_order(conn, order_number)
    log_activity(conn, username, full_name, 'Approved order', order_number=order_number)
    conn.close()
    flash(f'Order {order_number} approved and marked ready.', 'success')
    return redirect(url_for('warehouse'))


@app.route('/warehouse/order/<order_number>/reject', methods=['POST'])
@login_required
def warehouse_order_reject(order_number):
    username, full_name = _session_user()
    conn = get_conn()
    reject_warehouse_order(conn, order_number)
    log_activity(conn, username, full_name, 'Pushed back order (returned to In Prep)', order_number=order_number)
    conn.close()
    flash(f'Order {order_number} sent back to In Prep.', 'warning')
    return redirect(url_for('warehouse'))


@app.route('/admin/activity')
@admin_required
def admin_activity():
    conn    = get_conn()
    entries = get_activity_log(conn)
    conn.close()
    return render_template('admin_activity.html', entries=entries)


@app.route('/warehouse/order/<order_number>')
@login_required
def warehouse_order_summary(order_number):
    conn = get_conn()
    units, progress, total_items, photos = get_order_prep_summary(conn, order_number)
    conn.close()
    if not units:
        flash('Order not found.', 'error')
        return redirect(url_for('warehouse'))
    return render_template('warehouse_order.html',
                           order_number=order_number,
                           units=units,
                           progress=progress,
                           total_items=total_items,
                           photos=photos,
                           checklist_items=CHECKLIST_ITEMS)


# ── User management (admin only) ───────────────────────────────────────────────

@app.route('/admin/users')
@admin_required
def admin_users():
    conn  = get_conn()
    users = get_all_users(conn)
    conn.close()
    return render_template('admin_users.html', users=users)


@app.route('/admin/users/create', methods=['POST'])
@admin_required
def admin_users_create():
    username  = request.form.get('username', '').strip()
    password  = request.form.get('password', '')
    role      = request.form.get('role', 'warehouse')
    full_name = request.form.get('full_name', '').strip()
    if not username or not password:
        flash('Username and password are required.', 'error')
        return redirect(url_for('admin_users'))
    conn = get_conn()
    try:
        create_user(conn, username, generate_password_hash(password, method='pbkdf2:sha256'), role, full_name)
        flash(f'User "{username}" created successfully.', 'success')
    except Exception as e:
        flash(f'Error creating user: {e}', 'error')
    conn.close()
    return redirect(url_for('admin_users'))


@app.route('/admin/users/<int:user_id>/password', methods=['POST'])
@admin_required
def admin_users_password(user_id):
    password = request.form.get('password', '')
    if not password:
        flash('Password cannot be empty.', 'error')
        return redirect(url_for('admin_users'))
    conn = get_conn()
    update_user_password(conn, user_id, generate_password_hash(password, method='pbkdf2:sha256'))
    conn.close()
    flash('Password updated.', 'success')
    return redirect(url_for('admin_users'))


@app.route('/admin/users/<int:user_id>/delete', methods=['POST'])
@admin_required
def admin_users_delete(user_id):
    conn = get_conn()
    delete_user(conn, user_id)
    conn.close()
    flash('User deleted.', 'success')
    return redirect(url_for('admin_users'))


@app.route('/containers')
@login_required
def containers():
    conn = get_conn()
    manifests = list_manifests()
    container_ids = get_all_containers_in_sales(conn)
    containers_with_counts = []
    for cid in container_ids:
        allocated = len(get_sales_by_container(conn, cid))
        pre_sale  = conn.execute(
            "SELECT COUNT(*) AS n FROM units WHERE container_id=? AND status LIKE 'Pre-Sale%'",
            (cid,)
        ).fetchone()['n']
        containers_with_counts.append({
            'container_id': cid,
            'allocated':    allocated,
            'pre_sale':     pre_sale,
            'total':        allocated + pre_sale,
        })
    conn.close()
    return render_template('containers.html',
                           manifests=manifests,
                           containers=containers_with_counts)


@app.route('/containers/upload', methods=['POST'])
@login_required
def containers_upload():
    container_id = request.form.get('container_id_input', '').strip().upper()
    file = request.files.get('manifest_upload')

    if not container_id or not container_id.startswith('CNT-'):
        flash('Invalid container ID. Use the format CNT-### (e.g. CNT-005).', 'error')
        return redirect(url_for('containers'))

    if not file or file.filename == '':
        flash('No file selected.', 'error')
        return redirect(url_for('containers'))

    if not file.filename.lower().endswith('.csv'):
        flash('Only CSV files are accepted.', 'error')
        return redirect(url_for('containers'))

    os.makedirs(MANIFEST_DIR, exist_ok=True)
    dest_filename = f'container_manifest_{container_id}.csv'
    dest_path = os.path.join(MANIFEST_DIR, dest_filename)

    if os.path.exists(dest_path):
        flash(f'A manifest for {container_id} already exists ({dest_filename}). Delete or rename it first.', 'error')
        return redirect(url_for('containers'))

    file.save(dest_path)
    flash(f'Manifest uploaded as {dest_filename}. Select it below to parse.', 'success')
    return redirect(url_for('containers'))


@app.route('/containers/parse', methods=['POST'])
@login_required
def containers_parse():
    manifest_file = request.form.get('manifest_file')
    if not manifest_file:
        flash('No manifest file selected.', 'error')
        return redirect(url_for('containers'))

    if not os.path.isfile(manifest_file):
        flash(f'Manifest file not found: {manifest_file}', 'error')
        return redirect(url_for('containers'))

    try:
        container_id = container_id_from_filename(manifest_file)
        manifest_rows = load_manifest(manifest_file)
        conn = get_conn()
        n_matched, n_new = process_manifest(conn, manifest_rows, container_id)
        export_excel(conn)
        conn.close()
        flash(f'Container {container_id} parsed: {n_matched} matched variant(s), {n_new} new variant(s). {len(manifest_rows)} line items processed.', 'success')
    except Exception as e:
        flash(f'Error parsing manifest: {e}', 'error')

    return redirect(url_for('inventory'))


@app.route('/containers/receive', methods=['POST'])
@login_required
def containers_receive():
    container_id = request.form.get('container_id')
    if not container_id:
        flash('No container selected.', 'error')
        return redirect(url_for('containers'))

    try:
        conn = get_conn()
        date_arrived = datetime.date.today().isoformat()

        # Flip Pre-Sale units for this container → In Stock first
        result = conn.execute(
            "UPDATE units SET status = 'In Stock', date_received = ? WHERE container_id = ? AND status LIKE 'Pre-Sale%'",
            (date_arrived, container_id)
        )
        in_stock_count = result.rowcount

        # Clear awaiting_presale on sales_orders rows whose unit is now In Stock
        conn.execute("""
            UPDATE sales_orders SET awaiting_presale=0
            WHERE container_id=? AND awaiting_presale=1
        """, (container_id,))

        # Move allocated units → Warehouse, but only for orders where ALL units are now ready
        sales = get_sales_by_container(conn, container_id)
        moved_orders = set()
        for sale in sales:
            order_num = sale['order_number']
            if order_num in moved_orders:
                continue
            # Check if every unit's container has physically arrived
            still_pending = conn.execute("""
                SELECT COUNT(*) FROM sales_orders so
                WHERE so.order_number=?
                  AND NOT EXISTS (
                      SELECT 1 FROM units u
                      WHERE u.container_id = so.container_id
                        AND u.status = 'In Stock'
                  )
            """, (order_num,)).fetchone()[0]
            if still_pending == 0:
                # All units ready — move entire order to warehouse
                all_rows = conn.execute(
                    "SELECT * FROM sales_orders WHERE order_number=?", (order_num,)
                ).fetchall()
                for row in all_rows:
                    move_to_warehouse(conn, row['id'], date_arrived)
                moved_orders.add(order_num)

        conn.commit()

        if not sales and in_stock_count == 0:
            flash(f'No units found for container {container_id}.', 'error')
            conn.close()
            return redirect(url_for('containers'))

        export_excel(conn)
        conn.close()

        parts = []
        if sales:
            parts.append(f'{len(sales)} allocated unit(s) moved to warehouse')
        if in_stock_count:
            parts.append(f'{in_stock_count} Pre-Sale unit(s) moved to In Stock')
        flash(f'Container {container_id} received: {", ".join(parts)}.', 'success')
    except Exception as e:
        flash(f'Error receiving container: {e}', 'error')

    return redirect(url_for('warehouse'))


@app.route('/containers/cancel', methods=['POST'])
@login_required
def containers_cancel():
    order_number = request.form.get('order_number')
    source       = request.form.get('source', 'sales')
    serial       = request.form.get('serial')

    if not order_number:
        flash('No order number provided.', 'error')
        return redirect(url_for('sales'))

    try:
        conn = get_conn()
        username  = session.get('username', 'unknown')
        full_name = session.get('full_name', username)
        if source == 'warehouse' and serial:
            cancel_warehouse_order(conn, order_number, serial)
            log_activity(conn, username, full_name, 'Cancelled order (warehouse)',
                         detail=f'Serial {serial}', order_number=order_number, serial_number=serial)
            flash(f'Order {order_number} cancelled. Unit returned to In Stock.', 'success')
        else:
            count = cancel_order(conn, order_number)
            log_activity(conn, username, full_name, 'Cancelled order',
                         detail=f'{count} unit(s) returned to inventory', order_number=order_number)
            flash(f'Order {order_number} cancelled. {count} unit(s) returned to inventory.', 'success')
        export_excel(conn)
        conn.close()
    except Exception as e:
        flash(f'Error cancelling order: {e}', 'error')

    return redirect(url_for('sales'))


@app.route('/sales/available-units')
@login_required
def sales_available_units():
    conn  = get_conn()
    units = [dict(u) for u in get_available_units(conn)]
    conn.close()
    return jsonify(units)


@app.route('/sales/change-order', methods=['POST'])
@login_required
def sales_change_order():
    order_number = request.form.get('order_number')
    old_serial   = request.form.get('old_serial', '').strip()
    new_unit_id  = request.form.get('new_unit_id', '').strip()
    source       = request.form.get('source', 'sales')
    cr_id        = request.form.get('cr_id', '')
    co_action    = request.form.get('co_action', 'swap')

    if not order_number:
        flash('Missing order number.', 'error')
        return redirect(url_for('sales'))

    try:
        from datetime import datetime
        username  = session.get('username', 'unknown')
        full_name = session.get('full_name', username)
        conn      = get_conn()
        swap_ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if co_action == 'swap':
            if not old_serial or not new_unit_id:
                flash('Missing fields for swap.', 'error')
                return redirect(url_for('change_request_detail', cr_id=cr_id) if cr_id else url_for('sales'))

            old_info = conn.execute("""
                SELECT u.serial_number, v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku
                FROM units u JOIN variants v ON v.id=u.variant_id WHERE u.serial_number=?
            """, (old_serial,)).fetchone()
            new_unit = conn.execute("""
                SELECT u.id, u.serial_number, u.status, v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku
                FROM units u JOIN variants v ON v.id=u.variant_id WHERE u.id=?
            """, (int(new_unit_id),)).fetchone()
            new_serial = new_unit['serial_number'] if new_unit else new_unit_id
            is_presale = new_unit and new_unit['status'].startswith('Pre-Sale')

            change_order_unit(conn, order_number, old_serial, int(new_unit_id), source=source)
            if cr_id:
                if old_info:
                    log_cr_unit_change(conn, int(cr_id), 'removed', dict(old_info), swap_ts=swap_ts)
                if new_unit:
                    log_cr_unit_change(conn, int(cr_id), 'added', dict(new_unit), swap_ts=swap_ts)

            action = 'Swapped unit (pre-sale — order moved to Allocated)' if is_presale else 'Changed order unit'
            log_activity(conn, username, full_name, action,
                         detail=f'{old_serial} → {new_serial}',
                         order_number=order_number, serial_number=old_serial)
            if is_presale:
                flash(f'Order {order_number}: swapped to pre-sale — order moved back to Allocated.', 'warning')
            else:
                flash(f'Order {order_number}: unit swapped successfully.', 'success')

        elif co_action == 'remove':
            if not old_serial:
                flash('No unit selected to remove.', 'error')
                return redirect(url_for('change_request_detail', cr_id=cr_id) if cr_id else url_for('sales'))

            old_info = conn.execute("""
                SELECT u.serial_number, v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku
                FROM units u JOIN variants v ON v.id=u.variant_id WHERE u.serial_number=?
            """, (old_serial,)).fetchone()
            remove_unit_from_order(conn, order_number, old_serial, source)
            if cr_id and old_info:
                log_cr_unit_change(conn, int(cr_id), 'removed', dict(old_info), swap_ts=swap_ts)
            log_activity(conn, username, full_name, 'Removed unit from order',
                         detail=old_serial, order_number=order_number)
            flash(f'Order {order_number}: unit {old_serial} removed.', 'success')

        elif co_action == 'add':
            if not new_unit_id:
                flash('No unit selected to add.', 'error')
                return redirect(url_for('change_request_detail', cr_id=cr_id) if cr_id else url_for('sales'))

            customer = request.form.get('customer', '')
            added = add_unit_to_order(conn, order_number, customer, int(new_unit_id), source)
            if cr_id:
                log_cr_unit_change(conn, int(cr_id), 'added', added, swap_ts=swap_ts)
            log_activity(conn, username, full_name, 'Added unit to order',
                         detail=added.get('serial_number'), order_number=order_number)
            flash(f'Order {order_number}: unit {added.get("serial_number")} added.', 'success')

        export_excel(conn)
        conn.close()

    except Exception as e:
        flash(f'Error: {e}', 'error')

    if cr_id:
        return redirect(url_for('change_request_detail', cr_id=cr_id))
    if source == 'warehouse':
        return redirect(url_for('warehouse'))
    return redirect(url_for('sales'))


@app.route('/purchase-orders')
@login_required
def purchase_orders():
    conn = get_conn()
    negative_variants = get_variants_with_negative_variance(conn)
    conn.close()
    po_files = list_po_files()
    return render_template('purchase_orders.html',
                           negative_variants=negative_variants,
                           po_files=po_files)


@app.route('/purchase-orders/generate', methods=['POST'])
@login_required
def purchase_orders_generate():
    try:
        conn = get_conn()
        negative_variants = get_variants_with_negative_variance(conn)

        if not negative_variants:
            flash('No variants with negative variance. No PO generated.', 'info')
            conn.close()
            return redirect(url_for('purchase_orders'))

        os.makedirs(PO_DIR, exist_ok=True)
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'purchase_order_{timestamp}.csv'
        filepath = os.path.join(PO_DIR, filename)

        with open(filepath, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Design Name', 'Size', 'Finish', 'Swing', 'Glass Type', 'SKU', 'Qty Needed'])
            for v in negative_variants:
                qty_needed = v['optimal_count'] - (v['in_stock'] + v['in_prod'] + v['pre_sale'])
                writer.writerow([
                    v['design_name'], v['size'], v['finish'], v['swing'], v['glass_type'],
                    v['sku'], qty_needed
                ])
                # Seed In Production units
                for _ in range(qty_needed):
                    add_unit(conn, v['id'], None, 'In Production')

        export_excel(conn)
        conn.close()
        flash(f'Purchase order generated: {filename} ({len(negative_variants)} variant(s), In Production units seeded).', 'success')
    except Exception as e:
        flash(f'Error generating purchase order: {e}', 'error')

    return redirect(url_for('purchase_orders'))


@app.route('/optimal-review')
@login_required
def optimal_review():
    import math
    conn   = get_conn()
    rows   = get_variants_for_review(conn)
    today  = datetime.date.today()
    conn.close()

    variants = []
    for r in rows:
        total_sold      = r['total_sold'] or 0
        first_sale_date = r['first_sale_date']

        if total_sold > 0 and first_sale_date:
            days_of_data = max((today - datetime.date.fromisoformat(first_sale_date)).days, 1)

            if days_of_data >= 90:
                confidence       = 'High'
                quarterly_demand = (total_sold / days_of_data) * 90
                suggested        = max(math.ceil(quarterly_demand * 1.5), 1)
            elif days_of_data >= 30:
                confidence       = 'Medium'
                quarterly_demand = (total_sold / days_of_data) * 90
                suggested        = max(math.ceil(quarterly_demand * 1.2), 1)
            else:
                # Too little data to project — just use what actually sold + 1 buffer
                confidence = 'Low'
                suggested  = total_sold + 1
        else:
            suggested  = r['optimal_count'] or 0
            confidence = 'No data'

        variants.append({
            'id':            r['id'],
            'design_name':   r['design_name'],
            'size':          r['size'],
            'finish':        r['finish'],
            'swing':         r['swing'],
            'glass_type':    r['glass_type'],
            'sku':           r['sku'],
            'in_stock':      r['in_stock'],
            'in_prod':       r['in_prod'],
            'pre_sale':      r['pre_sale'],
            'total_sold':    total_sold,
            'optimal_count': r['optimal_count'] or 0,
            'suggested':     suggested,
            'confidence':    confidence,
        })

    return render_template('optimal_review.html', variants=variants, today=today)


@app.route('/optimal-review/save', methods=['POST'])
@login_required
def optimal_review_save():
    conn = get_conn()
    updated = 0
    for key, value in request.form.items():
        if key.startswith('optimal_'):
            variant_id = int(key.split('_')[1])
            try:
                new_optimal = int(value)
                if new_optimal >= 0:
                    set_optimal_count(conn, variant_id, new_optimal)
                    updated += 1
            except (ValueError, TypeError):
                pass
    export_excel(conn)
    conn.close()
    flash(f'Optimal counts updated for {updated} variant(s).', 'success')
    return redirect(url_for('optimal_review'))


@app.route('/purchase-orders/view/<filename>')
@login_required
def purchase_orders_view(filename):
    filepath = os.path.join(PO_DIR, filename)
    if not os.path.isfile(filepath):
        flash(f'File not found: {filename}', 'error')
        return redirect(url_for('purchase_orders'))

    lines = []
    grand_cost = grand_retail = 0
    with open(filepath, newline='') as f:
        for row in csv.DictReader(f):
            qty        = int(row.get('Qty Needed', 1))
            unit_cost, retail = calculate_cost(row['Design Name'], row['Size'])
            line_cost   = round(unit_cost * qty, 2)
            line_retail = round(retail * qty, 2)
            grand_cost  += line_cost
            grand_retail += line_retail
            lines.append({
                'design':      row['Design Name'],
                'size':        row['Size'],
                'finish':      row['Finish'],
                'swing':       row['Swing'],
                'glass_type':  row['Glass Type'],
                'sku':         row.get('SKU', ''),
                'qty':         qty,
                'unit_cost':   unit_cost,
                'retail':      retail,
                'line_cost':   line_cost,
                'line_retail': line_retail,
            })

    return render_template('purchase_order_detail.html',
                           filename=filename,
                           lines=lines,
                           grand_cost=round(grand_cost, 2),
                           grand_retail=round(grand_retail, 2))


@app.route('/purchase-orders/download/<filename>')
@login_required
def purchase_orders_download(filename):
    return send_from_directory(PO_DIR, filename, as_attachment=True)


# ── Change Requests ────────────────────────────────────────────────────────────

@app.route('/change-requests')
@admin_required
def change_requests():
    conn = get_conn()
    requests = get_all_change_requests(conn)
    conn.close()
    return render_template('change_orders.html',
                           requests=requests, scenarios=SCENARIOS)


@app.route('/change-requests/new')
@admin_required
def change_request_new():
    order_number = request.args.get('order', '').strip()
    customer     = request.args.get('customer', '').strip()

    auto_scenario  = None
    needs_question = None   # 'pickup' | 'delivery' | None
    order_info     = None

    order_units     = []
    available_units = []

    if order_number:
        conn = get_conn()
        # Redirect to existing open CO for this order if one exists
        existing_cr = conn.execute(
            "SELECT id FROM change_requests WHERE order_number=? AND status != 'Resolved' ORDER BY id DESC LIMIT 1",
            (order_number,)
        ).fetchone()
        if existing_cr:
            conn.close()
            return redirect(url_for('change_request_detail', cr_id=existing_cr['id']))

        so_row = conn.execute(
            "SELECT status FROM sales_orders WHERE order_number=? LIMIT 1",
            (order_number,)
        ).fetchone()
        wh_row = conn.execute(
            "SELECT status, fulfillment_type FROM warehouse WHERE order_number=? LIMIT 1",
            (order_number,)
        ).fetchone()
        order_units     = [dict(r) for r in get_units_on_order(conn, order_number)]
        available_units = [dict(r) for r in get_available_units(conn)]
        conn.close()

        if so_row:
            auto_scenario = '1A'
            order_info = {'label': f'Allocated · {so_row["status"]}'}
        elif wh_row:
            status = wh_row['status']
            order_info = {'label': f'Warehouse · {status}'}
            if status in ('In Prep', 'Pending Review'):
                auto_scenario = '1A'
            elif status == 'Ready for Pickup':
                needs_question = 'pickup'
            elif status == 'Ready for Delivery':
                needs_question = 'delivery'
            else:
                auto_scenario = '1A'

    return render_template('change_order_new.html',
                           order_number=order_number, customer=customer,
                           order_info=order_info, auto_scenario=auto_scenario,
                           needs_question=needs_question, scenarios=SCENARIOS,
                           order_units=order_units, available_units=available_units)


@app.route('/change-requests/create', methods=['POST'])
@admin_required
def change_request_create():
    order_number   = request.form.get('order_number', '').strip()
    customer       = request.form.get('customer', '').strip()
    auto_scenario  = request.form.get('auto_scenario', '').strip()
    request_detail = request.form.get('request_detail', '').strip()
    notes          = request.form.get('notes', '').strip()
    username       = session.get('username', 'unknown')
    full_name      = session.get('full_name', username)

    if not order_number:
        flash('Order number is required.', 'error')
        return redirect(url_for('change_request_new'))

    if auto_scenario:
        scenario_id = auto_scenario
    else:
        # Determine from follow-up questions (Ready for Pickup/Delivery edge cases)
        pickup_departed   = request.form.get('pickup_departed', '')
        shipment_departed = request.form.get('shipment_departed', '')
        carrier_type      = request.form.get('carrier_type', '')

        if pickup_departed == 'no':
            scenario_id = '1A'
        elif pickup_departed == 'yes':
            scenario_id = '1D'
        elif shipment_departed == 'no':
            scenario_id = '1A'
        elif carrier_type == 'ltl':
            scenario_id = '1B'
        elif carrier_type == 'non-ltl':
            scenario_id = '1C'
        else:
            flash('Could not determine scenario — please answer all questions.', 'error')
            return redirect(url_for('change_request_new', order=order_number, customer=customer))

    if scenario_id not in SCENARIOS:
        flash('Invalid scenario.', 'error')
        return redirect(url_for('change_request_new', order=order_number, customer=customer))

    co_action   = request.form.get('co_action', '').strip()   # 'swap' | 'remove' | 'add'
    old_serial  = request.form.get('old_serial', '').strip()
    new_unit_id = request.form.get('new_unit_id', '').strip()
    swap_source = request.form.get('swap_source', 'sales').strip()

    conn  = get_conn()
    cr_id = create_change_request(conn, order_number, customer, 'stock', scenario_id,
                                  request_detail, notes, username)
    log_activity(conn, username, full_name, 'Created change order',
                 detail=f'Scenario {scenario_id}: {SCENARIOS[scenario_id]["title"]}',
                 order_number=order_number)

    action_msg = ''
    try:
        if co_action == 'swap' and old_serial and new_unit_id:
            old_info = conn.execute("""
                SELECT u.serial_number, v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku
                FROM units u JOIN variants v ON v.id=u.variant_id WHERE u.serial_number=?
            """, (old_serial,)).fetchone()
            new_unit = conn.execute("""
                SELECT u.id, u.serial_number, u.status, v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku
                FROM units u JOIN variants v ON v.id=u.variant_id WHERE u.id=?
            """, (int(new_unit_id),)).fetchone()
            is_presale = new_unit and new_unit['status'].startswith('Pre-Sale')
            change_order_unit(conn, order_number, old_serial, int(new_unit_id), source=swap_source)
            if old_info:
                log_cr_unit_change(conn, cr_id, 'removed', dict(old_info))
            if new_unit:
                log_cr_unit_change(conn, cr_id, 'added', dict(new_unit))
            act = 'Swapped unit (pre-sale)' if is_presale else 'Swapped unit'
            log_activity(conn, username, full_name, act,
                         detail=f'{old_serial} → {new_unit["serial_number"] if new_unit else new_unit_id}',
                         order_number=order_number)
            action_msg = f' Swapped {old_serial} → {new_unit["serial_number"] if new_unit else new_unit_id}.'
            export_excel(conn)

        elif co_action == 'remove' and old_serial:
            old_info = conn.execute("""
                SELECT u.serial_number, v.design_name, v.size, v.finish, v.swing, v.glass_type, v.sku
                FROM units u JOIN variants v ON v.id=u.variant_id WHERE u.serial_number=?
            """, (old_serial,)).fetchone()
            remove_unit_from_order(conn, order_number, old_serial, swap_source)
            if old_info:
                log_cr_unit_change(conn, cr_id, 'removed', dict(old_info))
            log_activity(conn, username, full_name, 'Removed unit from order',
                         detail=old_serial, order_number=order_number)
            action_msg = f' Removed {old_serial} from order.'
            export_excel(conn)

        elif co_action == 'add' and new_unit_id:
            added = add_unit_to_order(conn, order_number, customer, int(new_unit_id), swap_source)
            log_cr_unit_change(conn, cr_id, 'added', added)
            log_activity(conn, username, full_name, 'Added unit to order',
                         detail=added.get('serial_number'), order_number=order_number)
            action_msg = f' Added {added.get("serial_number")} to order.'
            export_excel(conn)

    except Exception as e:
        action_msg = f' (Action failed: {e})'

    conn.close()
    flash(f'Change order created — scenario {scenario_id}.{action_msg}', 'success')
    return redirect(url_for('change_request_detail', cr_id=cr_id))


@app.route('/change-requests/<int:cr_id>')
@admin_required
def change_request_detail(cr_id):
    conn = get_conn()
    cr   = get_change_request(conn, cr_id)
    if not cr:
        conn.close()
        flash('Change order not found.', 'error')
        return redirect(url_for('change_requests'))

    # Fetch the order's current units (sales_orders or warehouse)
    order_units = conn.execute("""
        SELECT so.serial_number, v.design_name, v.size, v.finish, v.swing, v.glass_type,
               so.status, so.container_id, 'sales' AS source
        FROM sales_orders so JOIN variants v ON v.id = so.variant_id
        WHERE so.order_number = ?
        UNION ALL
        SELECT wh.serial_number, v.design_name, v.size, v.finish, v.swing, v.glass_type,
               wh.status, wh.container_id, 'warehouse' AS source
        FROM warehouse wh JOIN variants v ON v.id = wh.variant_id
        WHERE wh.order_number = ?
    """, (cr['order_number'], cr['order_number'])).fetchall()

    available    = [dict(u) for u in get_available_units(conn)]
    unit_changes = get_cr_unit_changes(conn, [cr_id]).get(cr_id, [])
    conn.close()

    scenario = get_scenario(cr['scenario_id'])
    friction_colors = FRICTION_COLOR.get(scenario.get('friction_level', 'low'), ('#F5F5F7', '#6E6E73'))
    return render_template('change_order_detail.html',
                           cr=cr, scenario=scenario,
                           friction_colors=friction_colors,
                           non_negotiables=NON_NEGOTIABLES,
                           order_units=order_units,
                           available_units=available,
                           unit_changes=unit_changes)


@app.route('/change-requests/<int:cr_id>/update', methods=['POST'])
@admin_required
def change_request_update(cr_id):
    status     = request.form.get('status', 'Open')
    resolution = request.form.get('resolution', '').strip()
    notes      = request.form.get('notes', '').strip()
    username   = session.get('username', 'unknown')
    full_name  = session.get('full_name', username)
    conn = get_conn()
    cr   = get_change_request(conn, cr_id)
    update_change_request(conn, cr_id, status, resolution, notes)
    log_activity(conn, username, full_name, f'Change order → {status}',
                 detail=resolution or None,
                 order_number=cr['order_number'] if cr else None)
    conn.close()
    flash(f'Change order updated — {status}.', 'success')
    return redirect(url_for('change_request_detail', cr_id=cr_id))


@app.route('/change-requests/<int:cr_id>/warehouse-ack', methods=['POST'])
@login_required
def change_request_warehouse_ack(cr_id):
    username, full_name = _session_user()
    conn = get_conn()
    cr   = get_change_request(conn, cr_id)
    ack_warehouse_change_request(conn, cr_id, username)
    log_activity(conn, username, full_name, 'Change order acknowledged (warehouse)',
                 order_number=cr['order_number'] if cr else None)
    conn.close()
    flash('Marked as adjusted — change order removed from warehouse alerts.', 'success')
    return redirect(url_for('warehouse'))


# ── Sales Pipeline ─────────────────────────────────────────────────────────────

STAGE_LABELS = {
    'presale':        'Pre-Sale',
    'allocated':      'Allocated',
    'in_prep':        'In Prep',
    'pending_review': 'Pending Review',
    'ready_pickup':   'Ready for Pickup',
    'ready_delivery': 'Ready for Delivery',
    'fulfilled':      'Fulfilled',
}
STAGE_ORDER = ['presale', 'allocated', 'in_prep', 'pending_review', 'ready_pickup', 'ready_delivery', 'fulfilled']


@app.route('/pipeline')
@admin_required
def pipeline():
    conn = get_conn()
    orders = get_pipeline_orders(conn)
    conn.close()
    return render_template('pipeline.html', orders=orders,
                           stage_labels=STAGE_LABELS, stage_order=STAGE_ORDER)


@app.route('/pipeline/<order_number>')
@admin_required
def pipeline_detail(order_number):
    conn = get_conn()
    detail = get_pipeline_order_detail(conn, order_number)
    # Sidebar: all orders in the same stage
    stage = request.args.get('stage') or detail.get('stage', '')
    all_orders = get_pipeline_orders(conn)
    sidebar_orders = [o for o in all_orders if o['pipeline_stage'] == stage]
    conn.close()
    return render_template('pipeline_detail.html', detail=detail, order_number=order_number,
                           sidebar_orders=sidebar_orders, sidebar_stage=stage,
                           stage_labels=STAGE_LABELS)


@app.route('/pipeline/<order_number>/details', methods=['POST'])
@admin_required
def pipeline_update_details(order_number):
    conn = get_conn()
    upsert_order_details(
        conn, order_number,
        request.form.get('phone', ''),
        request.form.get('email', ''),
        request.form.get('address', ''),
        request.form.get('city', ''),
        request.form.get('state', ''),
        request.form.get('zip', ''),
        request.form.get('notes', ''),
    )
    conn.close()
    flash('Contact info saved.', 'success')
    return redirect(url_for('pipeline_detail', order_number=order_number))


@app.route('/pipeline/<order_number>/photos', methods=['POST'])
@admin_required
def pipeline_upload_photo(order_number):
    caption  = request.form.get('caption', '')
    filename = _save_photo(request.files.get('photo'), os.path.join(ORDER_UPLOAD_DIR, order_number))
    if not filename:
        flash('Please upload a JPG, PNG, or WebP image.', 'error')
        return redirect(url_for('pipeline_detail', order_number=order_number))

    conn = get_conn()
    add_order_photo(conn, order_number, filename, caption, session.get('username'))
    conn.close()
    flash('Photo uploaded.', 'success')
    return redirect(url_for('pipeline_detail', order_number=order_number))


@app.route('/pipeline/photos/<int:photo_id>/delete', methods=['POST'])
@admin_required
def pipeline_delete_photo(photo_id):
    conn = get_conn()
    filename, order_number = delete_order_photo(conn, photo_id)
    conn.close()
    if filename and order_number:
        try:
            os.remove(os.path.join(ORDER_UPLOAD_DIR, order_number, filename))
        except OSError:
            pass
    return redirect(url_for('pipeline_detail', order_number=order_number) if order_number else url_for('pipeline'))


if __name__ == '__main__':
    from db import init_db
    init_db()
    conn = get_conn()
    seed_admin(conn, _LEGACY_ADMIN_USER, _LEGACY_ADMIN_HASH)
    # Randomly assign fulfillment_type to any rows that have none
    for table in ('warehouse', 'sales_orders'):
        rows = conn.execute(f"SELECT id FROM {table} WHERE fulfillment_type IS NULL OR fulfillment_type=''").fetchall()
        for row in rows:
            conn.execute(f"UPDATE {table} SET fulfillment_type=? WHERE id=?",
                         (random.choice(['pickup', 'delivery']), row['id']))
    conn.commit()
    conn.close()
    app.run(debug=True, port=5001)
