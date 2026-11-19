import io
import re
from datetime import date, timedelta

import pytest
from openpyxl import load_workbook

from app import create_app

ADMIN = ('admin@example.com', 'adminpass1')


@pytest.fixture()
def app(tmp_path):
    return create_app({
        'DATA_DIR': str(tmp_path), 'SECRET_KEY': 'test', 'TESTING': True,
        'ADMIN_EMAIL': ADMIN[0], 'ADMIN_PASSWORD': ADMIN[1],
    })


@pytest.fixture()
def client(app):
    return app.test_client()


def token(client, path=None):
    """Return the session CSRF token (works whether or not the user is logged in)."""
    with client.session_transaction() as sess:
        sess.setdefault('_csrf', 'testtoken')
        return sess['_csrf']


def test_forms_render_csrf_token(client):
    for path in ('/login', '/register'):
        html = client.get(path).get_data(as_text=True)
        assert re.search(r'name="_csrf" value="[0-9a-f]{32}"', html)


def post(client, path, data=None, page='/login', **kw):
    data = dict(data or {})
    data['_csrf'] = token(client, page)
    return client.post(path, data=data, **kw)


def login(client, email, password):
    return post(client, '/login', {'email': email, 'password': password})


def signup(client, email='p1@example.com', name='Player One', dob='2000-06-15'):
    return post(client, '/register', {
        'name': name, 'email': email, 'phone': '+46 70 123 4567', 'gender': 'Female',
        'dob': dob, 'password': 'password1', 'password2': 'password1'}, page='/register')


def iso(days):
    return (date.today() + timedelta(days=days)).isoformat()


def make_tournament(client, name='Spring Open', reg=5, cancel=3, start=10, end=11):
    r = post(client, '/admin/tournament/new', {
        'name': name, 'start_date': iso(start), 'end_date': iso(end),
        'registration_deadline': iso(reg), 'cancellation_deadline': iso(cancel),
        'cat_name': ['Singles', 'Doubles', 'Mixed Doubles'], 'cat_type': ['singles', 'doubles', 'mixed'],
    }, page='/admin/tournament/new')
    assert r.status_code == 302
    return int(r.headers['Location'].rstrip('/').split('/')[-1])


def category_ids(app, tid):
    import sqlite3
    conn = sqlite3.connect(app.config['DB_PATH'])
    rows = conn.execute('SELECT name, id FROM categories WHERE tournament_id = ?', (tid,)).fetchall()
    conn.close()
    return dict(rows)


def test_startup_is_idempotent_for_concurrent_workers(app, tmp_path):
    """Several gunicorn workers call create_app() on the same fresh DB."""
    for _ in range(3):
        create_app({'DATA_DIR': str(tmp_path), 'SECRET_KEY': 'test',
                    'ADMIN_EMAIL': ADMIN[0], 'ADMIN_PASSWORD': ADMIN[1]})
    import sqlite3
    conn = sqlite3.connect(app.config['DB_PATH'])
    assert conn.execute('SELECT COUNT(*) FROM users WHERE is_admin = 1').fetchone()[0] == 1
    conn.close()


def test_health(client):
    assert client.get('/api/health').get_json()['status'] == 'healthy'


def test_csrf_required(client):
    assert client.post('/login', data={'email': 'a@b.c', 'password': 'x'}).status_code == 400


def test_signup_validation_and_duplicate(client):
    bad = post(client, '/register', {'name': '', 'email': 'nope', 'phone': 'x', 'gender': '', 'dob': '2999-01-01',
                                     'password': 'short', 'password2': 'other'}, page='/register')
    assert bad.status_code == 200
    assert signup(client).status_code == 302
    post(client, '/logout')
    dup = signup(client)
    assert dup.status_code == 200 and b'already exists' in dup.data


def test_non_admin_forbidden_and_anonymous_redirected(client):
    assert client.get('/admin').status_code == 302
    signup(client)
    assert client.get('/admin').status_code == 403
    assert client.get('/admin/export.xlsx').status_code == 403


def test_admin_login_creates_tournament_with_categories(app, client):
    assert login(client, *ADMIN).status_code == 302
    tid = make_tournament(client)
    cats = category_ids(app, tid)
    assert set(cats) == {'Singles', 'Doubles', 'Mixed Doubles'}
    assert b'Spring Open' in client.get('/admin').data


def test_tournament_date_validation(client):
    login(client, *ADMIN)
    r = post(client, '/admin/tournament/new', {
        'name': 'Bad', 'start_date': iso(5), 'end_date': iso(4),
        'registration_deadline': iso(9), 'cancellation_deadline': iso(2)}, page='/admin/tournament/new')
    assert r.status_code == 200
    assert b'End date cannot be before' in r.data and b'Registration deadline must be' in r.data


def test_user_registration_flow_partner_optional_and_cancel(app, client):
    login(client, *ADMIN)
    tid = make_tournament(client)
    cats = category_ids(app, tid)
    post(client, '/logout')

    signup(client)
    page = f'/tournament/{tid}'
    # doubles with named partner, mixed open, singles ignores partner
    post(client, f'{page}/register', {'category_id': cats['Doubles'], 'partner_name': '  Anna   Svensson '}, page=page)
    post(client, f'{page}/register', {'category_id': cats['Mixed Doubles'], 'partner_name': ''}, page=page)
    post(client, f'{page}/register', {'category_id': cats['Singles'], 'partner_name': 'ignored'}, page=page)
    my = client.get('/my').get_data(as_text=True)
    assert 'Anna Svensson' in my and 'Open to any partner' in my and 'ignored' not in my

    dup = post(client, f'{page}/register', {'category_id': cats['Doubles']}, page=page, follow_redirects=True)
    assert b'already registered' in dup.data

    # cancel one
    import sqlite3
    conn = sqlite3.connect(app.config['DB_PATH'])
    rid = conn.execute('SELECT r.id FROM registrations r WHERE category_id = ?', (cats['Doubles'],)).fetchone()[0]
    conn.close()
    post(client, f'/registration/{rid}/cancel', page='/my')
    assert 'Anna Svensson' not in client.get('/my').get_data(as_text=True)


def test_registration_closed_after_deadline(app, client):
    login(client, *ADMIN)
    tid = make_tournament(client, name='Closed', reg=-1, cancel=-1, start=3, end=3)
    cats = category_ids(app, tid)
    post(client, '/logout')
    signup(client)
    r = post(client, f'/tournament/{tid}/register', {'category_id': cats['Singles']},
             page=f'/tournament/{tid}', follow_redirects=True)
    assert b'closed' in r.data.lower()
    assert b'Singles' not in client.get('/my').data


def test_cannot_cancel_after_cancellation_deadline(app, client):
    login(client, *ADMIN)
    tid = make_tournament(client, name='Locked', reg=5, cancel=3)
    cats = category_ids(app, tid)
    post(client, '/logout')
    signup(client)
    post(client, f'/tournament/{tid}/register', {'category_id': cats['Singles']}, page=f'/tournament/{tid}')
    import sqlite3
    conn = sqlite3.connect(app.config['DB_PATH'])
    conn.execute('UPDATE tournaments SET cancellation_deadline = ?', (iso(-1),))
    conn.commit()
    rid = conn.execute('SELECT id FROM registrations').fetchone()[0]
    conn.close()
    r = post(client, f'/registration/{rid}/cancel', page='/my', follow_redirects=True)
    assert b'deadline has passed' in r.data


def test_user_cannot_cancel_someone_elses_registration(app, client):
    login(client, *ADMIN)
    tid = make_tournament(client)
    cats = category_ids(app, tid)
    post(client, '/logout')
    signup(client, 'a@example.com', 'A')
    post(client, f'/tournament/{tid}/register', {'category_id': cats['Singles']}, page=f'/tournament/{tid}')
    post(client, '/logout')
    signup(client, 'b@example.com', 'B')
    assert post(client, '/registration/1/cancel', page='/my').status_code == 404


def test_excel_export(app, client):
    login(client, *ADMIN)
    tid = make_tournament(client, name='=Evil Cup')
    cats = category_ids(app, tid)
    post(client, '/logout')
    signup(client, 'x@example.com', '=HYPERLINK("http://evil")', dob='2000-06-15')
    post(client, f'/tournament/{tid}/register', {'category_id': cats['Doubles'], 'partner_name': 'Bob'},
         page=f'/tournament/{tid}')
    post(client, '/logout')
    login(client, *ADMIN)

    r = client.get(f'/admin/tournament/{tid}/export.xlsx')
    assert r.status_code == 200 and 'spreadsheetml' in r.mimetype
    ws = load_workbook(io.BytesIO(r.data))['Registrations']
    header = [c.value for c in ws[1]]
    row = dict(zip(header, [c.value for c in ws[2]]))
    assert row['Category'] == 'Doubles' and row['Partner'] == 'Bob' and row['Date of birth'] == '2000-06-15'
    start = date.today() + timedelta(days=10)
    assert row['Age at tournament start'] == start.year - 2000 - ((start.month, start.day) < (6, 15))
    assert row['Player'].startswith("'=") and row['Tournament'].startswith("'=")  # formula injection neutralised

    wb = load_workbook(io.BytesIO(client.get('/admin/export.xlsx').data))
    assert wb.sheetnames == ['Registrations', 'Users']


def test_admin_category_and_tournament_management(app, client):
    login(client, *ADMIN)
    tid = make_tournament(client)
    post(client, f'/admin/tournament/{tid}/categories', {'name': "Men's Singles", 'type': 'singles'},
         page=f'/admin/tournament/{tid}')
    assert "Men&#39;s Singles" in client.get(f'/admin/tournament/{tid}').get_data(as_text=True)
    cid = category_ids(app, tid)["Men's Singles"]
    post(client, f'/admin/category/{cid}/delete', page=f'/admin/tournament/{tid}')
    assert "Men's Singles" not in category_ids(app, tid)
    post(client, f'/admin/tournament/{tid}/edit', {
        'name': 'Renamed', 'start_date': iso(10), 'end_date': iso(11),
        'registration_deadline': iso(5), 'cancellation_deadline': iso(3)}, page=f'/admin/tournament/{tid}/edit')
    assert b'Renamed' in client.get(f'/admin/tournament/{tid}').data
    post(client, f'/admin/tournament/{tid}/delete', page=f'/admin/tournament/{tid}')
    assert client.get(f'/admin/tournament/{tid}').status_code == 404
