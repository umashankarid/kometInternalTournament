"""Komet Internal Tournament - tournament creation and registration app.

Admins create tournaments (with categories) and export registrations to Excel.
Users sign up, browse tournaments and register for categories.
"""
import io
import logging
import os
import re
import secrets
import sqlite3
import sys
import time
from datetime import date, datetime
from functools import wraps

from flask import (Flask, abort, flash, g, jsonify, redirect, render_template,
                   request, send_file, session, url_for)
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.security import check_password_hash, generate_password_hash

logging.basicConfig(level=logging.INFO, stream=sys.stdout,
                    format='[%(asctime)s] [%(levelname)s] %(message)s')
logger = logging.getLogger(__name__)

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    email TEXT NOT NULL UNIQUE COLLATE NOCASE,
    phone TEXT NOT NULL,
    gender TEXT NOT NULL,
    dob TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    is_admin INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS tournaments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    start_date TEXT NOT NULL,
    end_date TEXT NOT NULL,
    registration_deadline TEXT NOT NULL,
    cancellation_deadline TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS categories (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tournament_id INTEGER NOT NULL REFERENCES tournaments(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    type TEXT NOT NULL CHECK (type IN ('singles', 'doubles', 'mixed')),
    UNIQUE (tournament_id, name)
);

CREATE TABLE IF NOT EXISTS registrations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category_id INTEGER NOT NULL REFERENCES categories(id) ON DELETE CASCADE,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    partner_name TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    UNIQUE (category_id, user_id)
);

CREATE INDEX IF NOT EXISTS idx_categories_tournament ON categories(tournament_id);
CREATE INDEX IF NOT EXISTS idx_registrations_user ON registrations(user_id);
CREATE INDEX IF NOT EXISTS idx_registrations_category ON registrations(category_id);
"""

CATEGORY_TYPES = ('singles', 'doubles', 'mixed')
GENDERS = ('Male', 'Female', 'Other')
DEFAULT_CATEGORIES = (('Singles', 'singles'), ('Doubles', 'doubles'), ('Mixed Doubles', 'mixed'))

EMAIL_RE = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')
PHONE_RE = re.compile(r'^\+?[0-9][0-9 \-()]{5,19}$')

LOGIN_MAX_FAILS = 5
LOGIN_WINDOW_SECONDS = 300


def parse_date(value):
    try:
        return datetime.strptime((value or '').strip(), '%Y-%m-%d').date()
    except ValueError:
        return None


def age_on(dob_str, ref):
    dob = parse_date(dob_str)
    if not dob or not ref:
        return ''
    return ref.year - dob.year - ((ref.month, ref.day) < (dob.month, dob.day))


def now_iso():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def xl(value):
    """Neutralise spreadsheet formula injection for user supplied text."""
    if isinstance(value, str) and value[:1] in ('=', '+', '-', '@', '\t', '\r'):
        return "'" + value
    return value


def safe_filename(text):
    return re.sub(r'[^A-Za-z0-9._-]+', '_', text).strip('_') or 'tournament'


def create_app(test_config=None):
    app = Flask(__name__)
    app.config.update(
        DATA_DIR=os.environ.get('DATA_DIR', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')),
        SECRET_KEY=os.environ.get('SECRET_KEY', ''),
        ADMIN_EMAIL=os.environ.get('ADMIN_EMAIL', ''),
        ADMIN_PASSWORD=os.environ.get('ADMIN_PASSWORD', ''),
        ADMIN_NAME=os.environ.get('ADMIN_NAME', 'Admin'),
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE='Lax',
        SESSION_COOKIE_SECURE=os.environ.get('COOKIE_SECURE', '0').lower() in ('1', 'true', 'yes'),
        MAX_CONTENT_LENGTH=1024 * 1024,
    )
    if test_config:
        app.config.update(test_config)

    os.makedirs(app.config['DATA_DIR'], exist_ok=True)
    app.config['DB_PATH'] = os.path.join(app.config['DATA_DIR'], 'tournaments.db')

    # Persist a generated secret key on the data volume when none is configured,
    # so sessions survive restarts without shipping a hardcoded default.
    if not app.config['SECRET_KEY']:
        key_file = os.path.join(app.config['DATA_DIR'], 'secret_key')
        if os.path.exists(key_file):
            with open(key_file) as fh:
                app.config['SECRET_KEY'] = fh.read().strip()
        if not app.config['SECRET_KEY']:
            app.config['SECRET_KEY'] = secrets.token_hex(32)
            fd = os.open(key_file, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
            with os.fdopen(fd, 'w') as fh:
                fh.write(app.config['SECRET_KEY'])
            logger.warning('SECRET_KEY not set - generated one and stored it in %s', key_file)

    # Coolify / reverse proxy: trust one hop of X-Forwarded-* headers.
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

    failed_logins = {}

    # ------------------------------------------------------------------ DB
    def connect():
        conn = sqlite3.connect(app.config['DB_PATH'], timeout=30)
        conn.row_factory = sqlite3.Row
        conn.execute('PRAGMA foreign_keys=ON')
        return conn

    def get_db():
        if 'db' not in g:
            g.db = connect()
        return g.db

    @app.teardown_appcontext
    def close_db(_exc):
        db = g.pop('db', None)
        if db is not None:
            db.close()

    def init_db():
        conn = connect()
        try:
            conn.execute('PRAGMA journal_mode=WAL')
            conn.executescript(SCHEMA)
            email = app.config['ADMIN_EMAIL'].strip().lower()
            password = app.config['ADMIN_PASSWORD']
            if email and password:
                # Idempotent on purpose: gunicorn boots several workers at once and
                # each runs this. INSERT OR IGNORE avoids a check-then-insert race.
                cur = conn.execute(
                    'INSERT OR IGNORE INTO users (name, email, phone, gender, dob, password_hash, is_admin, created_at) '
                    'VALUES (?, ?, ?, ?, ?, ?, 1, ?)',
                    (app.config['ADMIN_NAME'], email, '-', 'Other', '1970-01-01',
                     generate_password_hash(password), now_iso()))
                if cur.rowcount:
                    logger.info('Seeded admin user %s', email)
                conn.execute('UPDATE users SET is_admin = 1 WHERE email = ?', (email,))
                conn.commit()
            else:
                logger.warning('ADMIN_EMAIL / ADMIN_PASSWORD not set - no admin account will be created')
        finally:
            conn.close()

    init_db()

    # ------------------------------------------------------- request hooks
    @app.before_request
    def load_user():
        g.user = None
        uid = session.get('user_id')
        if uid:
            g.user = get_db().execute('SELECT * FROM users WHERE id = ?', (uid,)).fetchone()
            if g.user is None:
                session.clear()

    @app.before_request
    def csrf_protect():
        if request.method == 'POST':
            token = session.get('_csrf')
            sent = request.form.get('_csrf', '')
            if not token or not secrets.compare_digest(token, sent):
                abort(400, 'Invalid or missing CSRF token. Reload the page and try again.')

    @app.after_request
    def security_headers(resp):
        resp.headers.setdefault('X-Content-Type-Options', 'nosniff')
        resp.headers.setdefault('X-Frame-Options', 'DENY')
        resp.headers.setdefault('Referrer-Policy', 'same-origin')
        return resp

    @app.context_processor
    def inject_globals():
        def csrf_token():
            if '_csrf' not in session:
                session['_csrf'] = secrets.token_hex(16)
            return session['_csrf']
        return {'csrf_token': csrf_token, 'current_user': g.get('user'), 'today': date.today().isoformat()}

    def login_required(view):
        @wraps(view)
        def wrapper(*args, **kwargs):
            if g.user is None:
                flash('Please log in first.', 'error')
                return redirect(url_for('login'))
            return view(*args, **kwargs)
        return wrapper

    def admin_required(view):
        @wraps(view)
        def wrapper(*args, **kwargs):
            if g.user is None:
                flash('Please log in first.', 'error')
                return redirect(url_for('login'))
            if not g.user['is_admin']:
                abort(403)
            return view(*args, **kwargs)
        return wrapper

    @app.errorhandler(400)
    @app.errorhandler(403)
    @app.errorhandler(404)
    def error_page(err):
        return render_template('error.html', code=err.code, message=err.description), err.code

    # ------------------------------------------------------------- helpers
    def tournament_view(row):
        d = dict(row)
        today = date.today()
        d['registration_open'] = today <= parse_date(row['registration_deadline'])
        d['cancellation_open'] = today <= parse_date(row['cancellation_deadline'])
        d['finished'] = today > parse_date(row['end_date'])
        return d

    def get_tournament_or_404(tid):
        row = get_db().execute('SELECT * FROM tournaments WHERE id = ?', (tid,)).fetchone()
        if row is None:
            abort(404, 'Tournament not found.')
        return tournament_view(row)

    def validate_tournament(form):
        errors = []
        name = form.get('name', '').strip()
        if not name or len(name) > 120:
            errors.append('Tournament name is required (max 120 characters).')
        labels = (('start_date', 'Start date'), ('end_date', 'End date'),
                  ('registration_deadline', 'Registration deadline'),
                  ('cancellation_deadline', 'Cancellation deadline'))
        dates = {}
        for key, label in labels:
            dates[key] = parse_date(form.get(key))
            if dates[key] is None:
                errors.append(f'{label} is required (YYYY-MM-DD).')
        if not errors:
            if dates['end_date'] < dates['start_date']:
                errors.append('End date cannot be before the start date.')
            if dates['registration_deadline'] > dates['start_date']:
                errors.append('Registration deadline must be on or before the start date.')
            if dates['cancellation_deadline'] > dates['start_date']:
                errors.append('Cancellation deadline must be on or before the start date.')
        data = {'name': name}
        data.update({k: (v.isoformat() if v else '') for k, v in dates.items()})
        return data, errors

    def registration_rows(where='', params=()):
        sql = (
            'SELECT r.id AS registration_id, r.partner_name, r.created_at AS registered_at, '
            't.id AS tournament_id, t.name AS tournament, t.start_date, '
            'c.id AS category_id, c.name AS category, c.type AS category_type, '
            'u.id AS user_id, u.name AS player, u.email, u.phone, u.gender, u.dob '
            'FROM registrations r '
            'JOIN categories c ON c.id = r.category_id '
            'JOIN tournaments t ON t.id = c.tournament_id '
            'JOIN users u ON u.id = r.user_id ' + where +
            ' ORDER BY t.start_date, t.name, c.name, u.name')
        return get_db().execute(sql, params).fetchall()

    def partner_text(row):
        if row['category_type'] == 'singles':
            return ''
        return row['partner_name'] or 'Open to any partner'

    def build_workbook(rows, include_users=False):
        wb = Workbook()
        ws = wb.active
        ws.title = 'Registrations'
        headers = ['Tournament', 'Tournament start', 'Category', 'Type', 'Player', 'Email', 'Phone',
                   'Gender', 'Date of birth', 'Age at tournament start', 'Partner', 'Registered at']
        ws.append(headers)
        for row in rows:
            start = parse_date(row['start_date'])
            ws.append([xl(row['tournament']), row['start_date'], xl(row['category']), row['category_type'],
                       xl(row['player']), xl(row['email']), xl(row['phone']), row['gender'], row['dob'],
                       age_on(row['dob'], start), xl(partner_text(row)), row['registered_at']])
        sheets = [(ws, headers)]
        if include_users:
            us = wb.create_sheet('Users')
            uheaders = ['Name', 'Email', 'Phone', 'Gender', 'Date of birth', 'Age today', 'Admin', 'Joined']
            us.append(uheaders)
            for u in get_db().execute('SELECT * FROM users ORDER BY name').fetchall():
                us.append([xl(u['name']), xl(u['email']), xl(u['phone']), u['gender'], u['dob'],
                           age_on(u['dob'], date.today()), 'Yes' if u['is_admin'] else '', u['created_at']])
            sheets.append((us, uheaders))
        for sheet, hdrs in sheets:
            for cell in sheet[1]:
                cell.font = Font(bold=True)
            sheet.freeze_panes = 'A2'
            for idx, title in enumerate(hdrs, start=1):
                longest = max([len(str(title))] + [len(str(c.value or '')) for c in sheet[get_column_letter(idx)][1:]])
                sheet.column_dimensions[get_column_letter(idx)].width = min(longest + 2, 50)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    def send_xlsx(buf, filename):
        return send_file(buf, as_attachment=True, download_name=filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # ---------------------------------------------------------- public/user
    @app.get('/api/health')
    def health():
        try:
            get_db().execute('SELECT 1').fetchone()
        except sqlite3.Error as exc:
            return jsonify({'status': 'unhealthy', 'error': str(exc)}), 500
        return jsonify({'status': 'healthy', 'service': 'kometInternalTournament'})

    @app.get('/')
    def index():
        rows = get_db().execute('SELECT * FROM tournaments ORDER BY start_date, name').fetchall()
        tournaments = []
        for row in rows:
            t = tournament_view(row)
            t['categories'] = get_db().execute(
                'SELECT * FROM categories WHERE tournament_id = ? ORDER BY name', (row['id'],)).fetchall()
            tournaments.append(t)
        upcoming = [t for t in tournaments if not t['finished']]
        past = [t for t in tournaments if t['finished']]
        return render_template('index.html', upcoming=upcoming, past=past)

    @app.route('/register', methods=['GET', 'POST'])
    def register():
        if g.user:
            return redirect(url_for('index'))
        form = request.form if request.method == 'POST' else {}
        if request.method == 'POST':
            name = form.get('name', '').strip()
            email = form.get('email', '').strip().lower()
            phone = form.get('phone', '').strip()
            gender = form.get('gender', '')
            dob = parse_date(form.get('dob'))
            password = form.get('password', '')
            errors = []
            if not name or len(name) > 100:
                errors.append('Name is required (max 100 characters).')
            if not EMAIL_RE.match(email) or len(email) > 150:
                errors.append('A valid email address is required.')
            if not PHONE_RE.match(phone):
                errors.append('A valid phone number is required (digits, spaces, +, -, brackets).')
            if gender not in GENDERS:
                errors.append('Please select a gender.')
            if dob is None or dob >= date.today() or dob.year < 1900:
                errors.append('A valid date of birth in the past is required.')
            if len(password) < 8:
                errors.append('Password must be at least 8 characters.')
            if password != form.get('password2', ''):
                errors.append('Passwords do not match.')
            if not errors:
                try:
                    db = get_db()
                    cur = db.execute(
                        'INSERT INTO users (name, email, phone, gender, dob, password_hash, created_at) '
                        'VALUES (?, ?, ?, ?, ?, ?, ?)',
                        (name, email, phone, gender, dob.isoformat(), generate_password_hash(password), now_iso()))
                    db.commit()
                except sqlite3.IntegrityError:
                    errors.append('An account with this email already exists.')
                else:
                    session.clear()
                    session['user_id'] = cur.lastrowid
                    flash('Account created. Welcome!', 'success')
                    return redirect(url_for('index'))
            for err in errors:
                flash(err, 'error')
        return render_template('register.html', form=form, genders=GENDERS)

    @app.route('/login', methods=['GET', 'POST'])
    def login():
        if g.user:
            return redirect(url_for('index'))
        if request.method == 'POST':
            email = request.form.get('email', '').strip().lower()
            key = (request.remote_addr, email)
            now = time.time()
            attempts = [t for t in failed_logins.get(key, []) if now - t < LOGIN_WINDOW_SECONDS]
            failed_logins[key] = attempts
            if len(attempts) >= LOGIN_MAX_FAILS:
                flash('Too many failed attempts. Try again in a few minutes.', 'error')
                return render_template('login.html'), 429
            user = get_db().execute('SELECT * FROM users WHERE email = ?', (email,)).fetchone()
            if user and check_password_hash(user['password_hash'], request.form.get('password', '')):
                failed_logins.pop(key, None)
                session.clear()
                session['user_id'] = user['id']
                return redirect(url_for('admin_home') if user['is_admin'] else url_for('index'))
            attempts.append(now)
            failed_logins[key] = attempts
            flash('Invalid email or password.', 'error')
        return render_template('login.html')

    @app.post('/logout')
    def logout():
        session.clear()
        return redirect(url_for('index'))

    @app.get('/tournament/<int:tid>')
    def tournament_detail(tid):
        t = get_tournament_or_404(tid)
        categories = get_db().execute(
            'SELECT c.*, (SELECT COUNT(*) FROM registrations r WHERE r.category_id = c.id) AS entries '
            'FROM categories c WHERE c.tournament_id = ? ORDER BY c.name', (tid,)).fetchall()
        mine = {}
        if g.user:
            for r in get_db().execute(
                    'SELECT r.* FROM registrations r JOIN categories c ON c.id = r.category_id '
                    'WHERE c.tournament_id = ? AND r.user_id = ?', (tid, g.user['id'])).fetchall():
                mine[r['category_id']] = r
        return render_template('tournament.html', t=t, categories=categories, mine=mine)

    @app.post('/tournament/<int:tid>/register')
    @login_required
    def register_category(tid):
        t = get_tournament_or_404(tid)
        if not t['registration_open']:
            flash('Registration for this tournament is closed.', 'error')
            return redirect(url_for('tournament_detail', tid=tid))
        cat = get_db().execute('SELECT * FROM categories WHERE id = ? AND tournament_id = ?',
                               (request.form.get('category_id', type=int), tid)).fetchone()
        if cat is None:
            abort(404, 'Category not found.')
        partner = ' '.join(request.form.get('partner_name', '').split())[:100]
        if cat['type'] == 'singles':
            partner = ''
        try:
            db = get_db()
            db.execute('INSERT INTO registrations (category_id, user_id, partner_name, created_at) '
                       'VALUES (?, ?, ?, ?)', (cat['id'], g.user['id'], partner, now_iso()))
            db.commit()
            flash(f'Registered for {cat["name"]}.', 'success')
        except sqlite3.IntegrityError:
            flash('You are already registered for this category.', 'error')
        return redirect(url_for('tournament_detail', tid=tid))

    @app.post('/registration/<int:rid>/cancel')
    @login_required
    def cancel_registration(rid):
        row = get_db().execute(
            'SELECT r.id, r.user_id, t.id AS tid, t.cancellation_deadline FROM registrations r '
            'JOIN categories c ON c.id = r.category_id JOIN tournaments t ON t.id = c.tournament_id '
            'WHERE r.id = ?', (rid,)).fetchone()
        if row is None or row['user_id'] != g.user['id']:
            abort(404, 'Registration not found.')
        if date.today() > parse_date(row['cancellation_deadline']):
            flash('The cancellation deadline has passed. Please contact the admin.', 'error')
        else:
            get_db().execute('DELETE FROM registrations WHERE id = ?', (rid,))
            get_db().commit()
            flash('Registration cancelled.', 'success')
        return redirect(request.form.get('back') == 'tournament'
                        and url_for('tournament_detail', tid=row['tid']) or url_for('my_registrations'))

    @app.get('/my')
    @login_required
    def my_registrations():
        rows = registration_rows('WHERE r.user_id = ?', (g.user['id'],))
        items = []
        for r in rows:
            t = tournament_view(get_db().execute('SELECT * FROM tournaments WHERE id = ?',
                                                 (r['tournament_id'],)).fetchone())
            items.append({'row': r, 't': t, 'partner': partner_text(r)})
        return render_template('my.html', items=items)

    # ---------------------------------------------------------------- admin
    @app.get('/admin')
    @admin_required
    def admin_home():
        rows = get_db().execute(
            'SELECT t.*, (SELECT COUNT(*) FROM categories c WHERE c.tournament_id = t.id) AS category_count, '
            '(SELECT COUNT(*) FROM registrations r JOIN categories c ON c.id = r.category_id '
            ' WHERE c.tournament_id = t.id) AS registration_count '
            'FROM tournaments t ORDER BY t.start_date DESC, t.name').fetchall()
        user_count = get_db().execute('SELECT COUNT(*) FROM users').fetchone()[0]
        return render_template('admin_home.html', tournaments=[tournament_view(r) | {
            'category_count': r['category_count'], 'registration_count': r['registration_count']} for r in rows],
            user_count=user_count)

    @app.route('/admin/tournament/new', methods=['GET', 'POST'])
    @admin_required
    def admin_new_tournament():
        form = request.form if request.method == 'POST' else {}
        if request.method == 'POST':
            data, errors = validate_tournament(form)
            if not errors:
                db = get_db()
                cur = db.execute(
                    'INSERT INTO tournaments (name, start_date, end_date, registration_deadline, '
                    'cancellation_deadline, created_at) VALUES (?, ?, ?, ?, ?, ?)',
                    (data['name'], data['start_date'], data['end_date'], data['registration_deadline'],
                     data['cancellation_deadline'], now_iso()))
                tid = cur.lastrowid
                # Categories can be entered right in the create form (parallel lists).
                names = request.form.getlist('cat_name')
                types = request.form.getlist('cat_type')
                added = 0
                for cname, ctype in zip(names, types):
                    cname = cname.strip()[:80]
                    if cname and ctype in CATEGORY_TYPES:
                        try:
                            db.execute('INSERT INTO categories (tournament_id, name, type) VALUES (?, ?, ?)',
                                       (tid, cname, ctype))
                            added += 1
                        except sqlite3.IntegrityError:
                            pass
                db.commit()
                flash(f'Tournament created with {added} categor{"y" if added == 1 else "ies"}.', 'success')
                return redirect(url_for('admin_tournament', tid=tid))
            for err in errors:
                flash(err, 'error')
        return render_template('admin_tournament_form.html', form=form, editing=False,
                               category_types=CATEGORY_TYPES, defaults=DEFAULT_CATEGORIES,
                               posted_cats=list(zip(request.form.getlist('cat_name'),
                                                    request.form.getlist('cat_type'))))

    @app.route('/admin/tournament/<int:tid>/edit', methods=['GET', 'POST'])
    @admin_required
    def admin_edit_tournament(tid):
        t = get_tournament_or_404(tid)
        form = request.form if request.method == 'POST' else t
        if request.method == 'POST':
            data, errors = validate_tournament(request.form)
            if not errors:
                get_db().execute(
                    'UPDATE tournaments SET name = ?, start_date = ?, end_date = ?, registration_deadline = ?, '
                    'cancellation_deadline = ? WHERE id = ?',
                    (data['name'], data['start_date'], data['end_date'], data['registration_deadline'],
                     data['cancellation_deadline'], tid))
                get_db().commit()
                flash('Tournament updated.', 'success')
                return redirect(url_for('admin_tournament', tid=tid))
            for err in errors:
                flash(err, 'error')
        return render_template('admin_tournament_form.html', form=form, editing=True, t=t,
                               category_types=CATEGORY_TYPES, defaults=(), posted_cats=[])

    @app.post('/admin/tournament/<int:tid>/delete')
    @admin_required
    def admin_delete_tournament(tid):
        get_tournament_or_404(tid)
        get_db().execute('DELETE FROM tournaments WHERE id = ?', (tid,))
        get_db().commit()
        flash('Tournament deleted.', 'success')
        return redirect(url_for('admin_home'))

    @app.get('/admin/tournament/<int:tid>')
    @admin_required
    def admin_tournament(tid):
        t = get_tournament_or_404(tid)
        categories = get_db().execute(
            'SELECT * FROM categories WHERE tournament_id = ? ORDER BY name', (tid,)).fetchall()
        regs = registration_rows('WHERE t.id = ?', (tid,))
        by_cat = {c['id']: [] for c in categories}
        for r in regs:
            by_cat[r['category_id']].append({
                'row': r, 'partner': partner_text(r), 'age': age_on(r['dob'], parse_date(r['start_date']))})
        return render_template('admin_tournament.html', t=t, categories=categories, by_cat=by_cat,
                               category_types=CATEGORY_TYPES, total=len(regs))

    @app.post('/admin/tournament/<int:tid>/categories')
    @admin_required
    def admin_add_category(tid):
        get_tournament_or_404(tid)
        db = get_db()
        if request.form.get('defaults'):
            added = 0
            for cname, ctype in DEFAULT_CATEGORIES:
                try:
                    db.execute('INSERT INTO categories (tournament_id, name, type) VALUES (?, ?, ?)',
                               (tid, cname, ctype))
                    added += 1
                except sqlite3.IntegrityError:
                    pass
            db.commit()
            flash(f'Added {added} default categor{"y" if added == 1 else "ies"}.', 'success')
        else:
            cname = request.form.get('name', '').strip()[:80]
            ctype = request.form.get('type', '')
            if not cname or ctype not in CATEGORY_TYPES:
                flash('Category name and a valid type are required.', 'error')
            else:
                try:
                    db.execute('INSERT INTO categories (tournament_id, name, type) VALUES (?, ?, ?)',
                               (tid, cname, ctype))
                    db.commit()
                    flash(f'Category "{cname}" added.', 'success')
                except sqlite3.IntegrityError:
                    flash('A category with that name already exists in this tournament.', 'error')
        return redirect(url_for('admin_tournament', tid=tid))

    @app.post('/admin/category/<int:cid>/delete')
    @admin_required
    def admin_delete_category(cid):
        cat = get_db().execute('SELECT * FROM categories WHERE id = ?', (cid,)).fetchone()
        if cat is None:
            abort(404, 'Category not found.')
        get_db().execute('DELETE FROM categories WHERE id = ?', (cid,))
        get_db().commit()
        flash(f'Category "{cat["name"]}" deleted.', 'success')
        return redirect(url_for('admin_tournament', tid=cat['tournament_id']))

    @app.post('/admin/registration/<int:rid>/delete')
    @admin_required
    def admin_delete_registration(rid):
        row = get_db().execute(
            'SELECT c.tournament_id FROM registrations r JOIN categories c ON c.id = r.category_id '
            'WHERE r.id = ?', (rid,)).fetchone()
        if row is None:
            abort(404, 'Registration not found.')
        get_db().execute('DELETE FROM registrations WHERE id = ?', (rid,))
        get_db().commit()
        flash('Registration removed.', 'success')
        return redirect(url_for('admin_tournament', tid=row['tournament_id']))

    @app.get('/admin/tournament/<int:tid>/export.xlsx')
    @admin_required
    def admin_export_tournament(tid):
        t = get_tournament_or_404(tid)
        buf = build_workbook(registration_rows('WHERE t.id = ?', (tid,)))
        return send_xlsx(buf, f'{safe_filename(t["name"])}_registrations.xlsx')

    @app.get('/admin/export.xlsx')
    @admin_required
    def admin_export_all():
        buf = build_workbook(registration_rows(), include_users=True)
        return send_xlsx(buf, 'komet_tournaments_all_registrations.xlsx')

    return app


if __name__ == '__main__':
    create_app().run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
