"""Local browser interface for the banking importer (Windows and macOS)."""
import csv
import os
import time
from datetime import timedelta
import io
import secrets
import subprocess
import sys
import tempfile
import threading
import webbrowser
from pathlib import Path
from flask import Flask, abort, render_template, request, send_file, session, redirect, url_for

HERE = Path(__file__).resolve().parent
app = Flask(__name__)
HOSTED = os.environ.get('BANKING_HOSTED') == '1' or bool(os.environ.get('RENDER'))
PASSWORD = os.environ.get('BANKING_PASSWORD')
if HOSTED and (not PASSWORD or not os.environ.get('BANKING_SESSION_SECRET')):
    raise RuntimeError('Hosted mode requires BANKING_PASSWORD and BANKING_SESSION_SECRET.')
from werkzeug.security import generate_password_hash, check_password_hash
PASSWORD_HASH = generate_password_hash(PASSWORD) if PASSWORD else None
app.secret_key = os.environ.get('BANKING_SESSION_SECRET') or secrets.token_hex(32)
app.config.update(SESSION_COOKIE_SECURE=HOSTED, SESSION_COOKIE_HTTPONLY=True,
                  PERMANENT_SESSION_LIFETIME=timedelta(hours=8))
login_failures = []

@app.before_request
def protect_pages():
    if PASSWORD_HASH and request.endpoint not in ('login', 'robots', 'health') and not session.get('authenticated'):
        return redirect(url_for('login'))

@app.after_request
def privacy_headers(response):
    response.headers['X-Robots-Tag'] = 'noindex, nofollow, noarchive'
    response.headers['Cache-Control'] = 'no-store'
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['Content-Security-Policy'] = "default-src 'self'; style-src 'self' 'unsafe-inline'; frame-ancestors 'none'; form-action 'self'"
    return response

@app.get('/robots.txt')
def robots():
    return 'User-agent: *\nDisallow: /\n', 200, {'Content-Type': 'text/plain'}

@app.get('/health')
def health():
    return 'OK', 200

@app.route('/login', methods=['GET', 'POST'])
def login():
    if not PASSWORD_HASH:
        return redirect(url_for('index'))
    session.setdefault('csrf', secrets.token_urlsafe(32))
    error = None
    if request.method == 'POST':
        if not secrets.compare_digest(request.form.get('csrf', ''), session['csrf']):
            abort(403)
        now = time.monotonic()
        login_failures[:] = [t for t in login_failures if now - t < 300]
        if len(login_failures) >= 5:
            return render_template('login.html', csrf=session['csrf'], error='Too many attempts. Wait five minutes and try again.'), 429
        if check_password_hash(PASSWORD_HASH, request.form.get('password', '')):
            session.clear()
            session['authenticated'] = True
            session['csrf'] = secrets.token_urlsafe(32)
            session.permanent = True
            login_failures.clear()
            return redirect(url_for('index'))
        login_failures.append(now)
        error = 'Incorrect password.'
    return render_template('login.html', csrf=session['csrf'], error=error)

@app.post('/logout')
def logout():
    if not secrets.compare_digest(request.form.get('csrf', ''), session.get('csrf', '')):
        abort(403)
    app.extensions.get('results', {}).pop(session.get('result'), None)
    session.clear()
    return redirect(url_for('login'))
app.config.update(MAX_CONTENT_LENGTH=20 * 1024 * 1024, SESSION_COOKIE_SAMESITE='Strict')

@app.route('/', methods=['GET', 'POST'])
def index():
    session.setdefault('csrf', secrets.token_urlsafe(32))
    result = None
    error = None
    if request.method == 'POST':
        if not secrets.compare_digest(request.form.get('csrf', ''), session['csrf']):
            abort(403)
        report = request.files.get('report')
        workbook = request.files.get('workbook')
        if not report or not workbook or not report.filename.lower().endswith('.pdf') or not workbook.filename.lower().endswith('.xlsx'):
            error = 'Choose a PDF report and an Excel .xlsx workbook.'
        else:
            with tempfile.TemporaryDirectory(prefix='banking-import-') as folder:
                root = Path(folder)
                pdf, xlsx = root / 'report.pdf', root / 'banking.xlsx'
                report.save(pdf)
                workbook.save(xlsx)
                try:
                    run = subprocess.run([sys.executable, str(HERE / 'banking_import.py'), str(pdf), str(xlsx), '--output-dir', str(root / 'output'), '--apply'], capture_output=True, text=True, timeout=120)
                    reviews = list((root / 'output').glob('review-*.csv'))
                    review = reviews[0].read_bytes() if reviews else None
                    entries = list(csv.DictReader(io.StringIO(review.decode('utf-8-sig')))) if review else []
                    balances = [e for e in entries if 'LEDGER BALANCE' in e['Status']]
                    notices = [e['Source'] for e in entries if e['Status'] in ('ROUNDING NOTICE', 'REVIEW REQUIRED')]
                    changes = [e for e in entries if e['Status'] == 'CONFIRMED']
                    files = list((root / 'output').glob('*.xlsx'))
                    # Browser session stores only a random key; files remain local in memory.
                    token = secrets.token_urlsafe(32)
                    store = app.extensions.setdefault('results', {})
                    owner = session.setdefault('owner', secrets.token_urlsafe(32))
                    previous = session.get('result')
                    if previous: store.pop(previous, None)
                    if len(store) >= 20: store.pop(next(iter(store)))
                    store[token] = {'owner': owner, 'review': review, 'workbook': files[0].read_bytes() if run.returncode == 0 and files else None}
                    session['result'] = token
                    result = dict(success=run.returncode == 0, balances=balances, notices=notices, changes=changes, review=review is not None,
                                  date=entries[0]['Date'] if entries else '', sheet=entries[0]['Sheet'] if entries else '')
                    if run.returncode:
                        error = run.stderr.strip().removeprefix('Import stopped: ') or 'The report could not be imported. Check the file format.'
                except subprocess.TimeoutExpired:
                    error = 'The import took too long. Please check the PDF and workbook, then try again.'
    return render_template('index.html', csrf=session['csrf'], result=result, error=error)

@app.get('/download/<kind>')
def download(kind):
    if kind not in ('workbook', 'review'): abort(404)
    saved = app.extensions.get('results', {}).get(session.get('result'))
    if not saved or saved['owner'] != session.get('owner') or saved[kind] is None: abort(404)
    name = 'banking-updated.xlsx' if kind == 'workbook' else 'banking-review.csv'
    return send_file(io.BytesIO(saved[kind]), as_attachment=True, download_name=name)

@app.errorhandler(413)
def too_large(error):
    return render_template('index.html', csrf=session.get('csrf', ''), result=None, error='Files are too large. Choose files totalling less than 20 MB.'), 413

if __name__ == '__main__':
    print('Open http://127.0.0.1:8765 in your browser. Press Ctrl+C here to stop.')
    threading.Timer(1, lambda: webbrowser.open('http://127.0.0.1:8765')).start()
    app.run(host='127.0.0.1', port=8765, debug=False)
