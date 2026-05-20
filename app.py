from flask import Flask, request, Response, render_template_string, jsonify, session, redirect, url_for
import os, json, shutil, tempfile, base64
from fifo_engine import run_fifo
from extract import extract
from db import init_db, save_data, load_data, get_setting, set_setting, init_settings

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "etpmex-fifo-secret-2024")

LOGO_B64 = "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAzMjAgODAiIHdpZHRoPSIzMjAiIGhlaWdodD0iODAiPgogIDxkZWZzPgogICAgPGxpbmVhckdyYWRpZW50IGlkPSJnMSIgeDE9IjAlIiB5MT0iMCUiIHgyPSIxMDAlIiB5Mj0iMCUiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAlIiBzdG9wLWNvbG9yPSIjYzAzOTJiIi8+CiAgICAgIDxzdG9wIG9mZnNldD0iMTAwJSIgc3RvcC1jb2xvcj0iI2U3NGMzYyIvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICAgIDxsaW5lYXJHcmFkaWVudCBpZD0iZzIiIHgxPSIwJSIgeTE9IjAlIiB4Mj0iMTAwJSIgeTI9IjAlIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwJSIgc3RvcC1jb2xvcj0iI2Q2ZTRmNyIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEwMCUiIHN0b3AtY29sb3I9IiNhZWNiZjAiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImczIiB4MT0iMCUiIHkxPSIwJSIgeDI9IjEwMCUiIHkyPSIwJSI+CiAgICAgIDxzdG9wIG9mZnNldD0iMCUiIHN0b3AtY29sb3I9IiNlNzRjM2MiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIxMDAlIiBzdG9wLWNvbG9yPSIjYzAzOTJiIi8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogIDwvZGVmcz4KICA8ZyB0cmFuc2Zvcm09InRyYW5zbGF0ZSg0LCAxMCkiPgogICAgPHJlY3QgeD0iNCIgeT0iMiIgIHdpZHRoPSI1OCIgaGVpZ2h0PSIxMyIgcng9IjYuNSIgZmlsbD0idXJsKCNnMSkiIHRyYW5zZm9ybT0icm90YXRlKC0zMiwgMzMsIDguNSkiLz4KICAgIDxyZWN0IHg9IjQiIHk9IjIyIiB3aWR0aD0iNTgiIGhlaWdodD0iMTMiIHJ4PSI2LjUiIGZpbGw9InVybCgjZzIpIiB0cmFuc2Zvcm09InJvdGF0ZSgtMzIsIDMzLCAyOC41KSIvPgogICAgPHJlY3QgeD0iNCIgeT0iNDIiIHdpZHRoPSI1OCIgaGVpZ2h0PSIxMyIgcng9IjYuNSIgZmlsbD0idXJsKCNnMykiIHRyYW5zZm9ybT0icm90YXRlKC0zMiwgMzMsIDQ4LjUpIi8+CiAgPC9nPgogIDx0ZXh0IHg9IjkyIiB5PSI1NSIKICAgICAgICBmb250LWZhbWlseT0iJ0FyaWFsIEJsYWNrJywgJ0ZyYW5rbGluIEdvdGhpYyBIZWF2eScsICdIZWx2ZXRpY2EgTmV1ZScsIEFyaWFsLCBzYW5zLXNlcmlmIgogICAgICAgIGZvbnQtd2VpZ2h0PSI5MDAiIGZvbnQtc2l6ZT0iNTQiIGZpbGw9IndoaXRlIiBsZXR0ZXItc3BhY2luZz0iLTEiPkVHQzwvdGV4dD4KPC9zdmc+"

try:
    init_db()
    init_settings()
except Exception as e:
    print(f"DB init warning: {e}")

USERS = {
    "EGC.ADMIN": {"password": "EGC$admin!2026", "role": "admin"},
    "EGC.LPP":   {"password": "LPP!EGC$26",     "role": "potential_investor"},
    "EGC.LPA":   {"password": "LPA!EGC$26",     "role": "partner_investor"},
}

# ── Auth ───────────────────────────────────────────────────────────────────────
LOGIN_HTML = """<!DOCTYPE html>
<html><head>
<meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
<title>EGC</title>
  <link rel="icon" type="image/png" href="data:image/png;base64,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">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{--bg:#0d1117;--surface:#161c26;--border:#263044;--accent:#4a7cdc;--text:#e6edf3;--muted:#6e7f96;--error:#f85149;--font:'Inter',sans-serif}
body{background:var(--bg);color:var(--text);font-family:var(--font);min-height:100vh;display:flex;align-items:center;justify-content:center;
background-image:radial-gradient(ellipse 70% 40% at 50% 0%,rgba(74,124,220,0.07) 0%,transparent 70%)}
.wrap{width:100%;max-width:380px;padding:1rem}
.logo{display:flex;justify-content:center;margin-bottom:2rem}
.logo img{height:64px;width:auto}
.card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:2rem}
h2{font-size:1.1rem;font-weight:600;margin-bottom:0.3rem}
p{font-size:0.8rem;color:var(--muted);margin-bottom:1.5rem}
label{display:block;font-size:0.75rem;color:var(--muted);margin-bottom:0.4rem;margin-top:1rem}
input{width:100%;background:#0d1117;border:1px solid var(--border);border-radius:6px;padding:0.65rem 0.85rem;
color:var(--text);font-family:var(--font);font-size:0.85rem;outline:none;transition:border-color 0.15s}
input:focus{border-color:var(--accent)}
.err{color:var(--error);font-size:0.78rem;margin-top:0.75rem}
button{width:100%;margin-top:1.25rem;background:var(--accent);color:#fff;border:none;border-radius:8px;
padding:0.75rem;font-family:var(--font);font-size:0.9rem;font-weight:600;cursor:pointer;transition:opacity 0.15s}
button:hover{opacity:0.88}
</style></head><body>
<div class="wrap">
  <div class="logo"><img src="data:image/svg+xml;base64,{{ logo }}" alt="EGC"></div>
  <div class="card">
    <h2>Sign in</h2>
    {% if error %}<div class="err">{{ error }}</div>{% endif %}
    <form method="POST">
      <label>Username</label>
      <input name="username" type="text" autocomplete="username" required>
      <label>Password</label>
      <input name="password" type="password" autocomplete="current-password" required>
      <button type="submit">Sign in</button>
    </form>
  </div>
</div>
</body></html>"""

# dashboard.html loaded per-request below

@app.route('/login', methods=['GET','POST'])
def login():
    error = None
    if request.method == 'POST':
        u = request.form.get('username','').strip()
        p = request.form.get('password','').strip()
        user = USERS.get(u)
        if user and user['password'] == p:
            session['user'] = u
            session['role'] = user['role']
            return redirect(url_for('dashboard'))
        error = "Invalid username or password"
    return render_template_string(LOGIN_HTML, logo=LOGO_B64, error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/')
def index():
    if 'user' not in session:
        return redirect(url_for('login'))
    return redirect(url_for('dashboard'))

@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('login'))
    dash_path = os.path.join(os.path.dirname(__file__), 'dashboard.html')
    with open(dash_path) as f:
        html = f.read()
    return render_template_string(html, logo=LOGO_B64,
                                  role=session['role'], user=session['user'])

@app.route('/api/data')
def api_data():
    if 'user' not in session:
        return jsonify(error='Unauthorized'), 401
    try:
        markup_usd = float(get_setting("markup_usd", "0.02"))
        data = load_data()
    except Exception as e:
        print(f"[load_data error] {e}", flush=True)
        return jsonify(None)
    if not data:
        return jsonify(None)
    role = session['role']
    if role == 'investor':
        return jsonify({
            "uploaded_at":    data["uploaded_at"],
            "filename":       data["filename"],
            "meta":           data["meta"],
            "overall_summary": data["overall_summary"],
            "investment":     data.get("investment", {}),
            "bol_tab":        data.get("bol_tab", {}),
            "overview_exp":   data.get("overview_exp", {}),
        })
    if role == 'partner_investor':
        return jsonify({
            "uploaded_at":    data["uploaded_at"],
            "filename":       data["filename"],
            "meta":           data["meta"],
            "overall_summary": data["overall_summary"],
            "inventory":      data["inventory"],
            "fifo_rows":      data["fifo_rows"],
            "investment":     data.get("investment", {}),
            "bol_tab":        data.get("bol_tab", {}),
            "overview_exp":   data.get("overview_exp", {}),
            "markup":         True,
            "markup_usd":     markup_usd,
        })
    if role == 'potential_investor':
        return jsonify({
            "uploaded_at":    data["uploaded_at"],
            "filename":       data["filename"],
        })
    return jsonify(data)  # admin gets everything

@app.route('/process', methods=['POST'])
def process():
    if 'user' not in session or session['role'] != 'admin':
        return jsonify(error='Unauthorized'), 401
    if 'file' not in request.files:
        return jsonify(error='No file uploaded'), 400
    f = request.files['file']
    if not f.filename.endswith('.xlsx'):
        return jsonify(error='Must be an .xlsx file'), 400

    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, 'input.xlsx')
        dst = os.path.join(tmp, 'output.xlsx')
        f.save(src)
        try:
            run_fifo(src, dst)
            result = extract(dst, src_path=src)
            if len(result) == 7:
                overall_summary, inventory, fifo_rows, meta, investment, bol_tab, overview_exp = result
            elif len(result) == 6:
                overall_summary, inventory, fifo_rows, meta, investment, bol_tab = result
                overview_exp = {}
            elif len(result) == 5:
                overall_summary, inventory, fifo_rows, meta, investment = result
                bol_tab = {}; overview_exp = {}
            else:
                overall_summary, inventory, fifo_rows, meta = result
                investment = {}; bol_tab = {}; overview_exp = {}
            save_data(f.filename, overall_summary, inventory, fifo_rows, meta, investment, bol_tab, overview_exp)
        except Exception as e:
            return jsonify(error=str(e)), 500
        with open(dst, 'rb') as fh:
            data = fh.read()

    return Response(data,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename=FIFO_Output.xlsx'})



@app.route('/api/settings', methods=['POST'])
def save_settings():
    if 'user' not in session or session['role'] != 'admin':
        return jsonify(error='Unauthorized'), 403
    data = request.get_json()
    val = data.get('markup_usd', '0.02')
    try:
        v = round(float(val), 4)
        if v < 0 or v > 10:
            raise ValueError
        set_setting('markup_usd', v)
        return jsonify(ok=True, markup_usd=v)
    except:
        return jsonify(error='Invalid value'), 400

@app.route('/api/settings', methods=['GET'])
def get_settings():
    if 'user' not in session or session['role'] != 'admin':
        return jsonify(error='Unauthorized'), 403
    return jsonify(markup_usd=float(get_setting('markup_usd', '0.02')))

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)
