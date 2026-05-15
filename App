from flask import Flask, request, send_file, render_template_string, jsonify
import io, os, shutil, tempfile
from fifo_engine import run_fifo
 
app = Flask(__name__)
 
HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ETP — FIFO Engine</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Syne:wght@700;800&display=swap" rel="stylesheet">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
 
  :root {
    --bg:       #0b0e13;
    --surface:  #13181f;
    --border:   #1f2830;
    --accent:   #00e5a0;
    --accent2:  #005c40;
    --text:     #e8edf2;
    --muted:    #5a6675;
    --danger:   #ff4d4d;
    --mono:     'DM Mono', monospace;
    --display:  'Syne', sans-serif;
  }
 
  body {
    background: var(--bg);
    color: var(--text);
    font-family: var(--mono);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 2rem;
    background-image:
      radial-gradient(ellipse 80% 50% at 50% -10%, rgba(0,229,160,0.08) 0%, transparent 70%),
      repeating-linear-gradient(0deg, transparent, transparent 39px, rgba(255,255,255,0.02) 39px, rgba(255,255,255,0.02) 40px),
      repeating-linear-gradient(90deg, transparent, transparent 39px, rgba(255,255,255,0.02) 39px, rgba(255,255,255,0.02) 40px);
  }
 
  .container {
    width: 100%;
    max-width: 580px;
  }
 
  .logo {
    font-family: var(--display);
    font-size: 0.7rem;
    font-weight: 800;
    letter-spacing: 0.3em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 0.4rem;
  }
 
  h1 {
    font-family: var(--display);
    font-size: 2.6rem;
    font-weight: 800;
    line-height: 1.05;
    letter-spacing: -0.02em;
    margin-bottom: 0.5rem;
  }
 
  h1 span {
    color: var(--accent);
  }
 
  .subtitle {
    color: var(--muted);
    font-size: 0.78rem;
    letter-spacing: 0.05em;
    margin-bottom: 2.5rem;
    line-height: 1.6;
  }
 
  .card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 2rem;
    position: relative;
    overflow: hidden;
  }
 
  .card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, var(--accent), transparent);
  }
 
  .drop-zone {
    border: 1.5px dashed var(--border);
    border-radius: 3px;
    padding: 2.5rem 1.5rem;
    text-align: center;
    cursor: pointer;
    transition: border-color 0.2s, background 0.2s;
    position: relative;
    margin-bottom: 1.5rem;
  }
 
  .drop-zone:hover, .drop-zone.dragover {
    border-color: var(--accent);
    background: rgba(0,229,160,0.03);
  }
 
  .drop-zone input[type=file] {
    position: absolute;
    inset: 0;
    opacity: 0;
    cursor: pointer;
    width: 100%;
    height: 100%;
  }
 
  .drop-icon {
    font-size: 2rem;
    margin-bottom: 0.75rem;
    display: block;
  }
 
  .drop-label {
    font-size: 0.8rem;
    color: var(--muted);
    line-height: 1.6;
  }
 
  .drop-label strong {
    color: var(--accent);
    font-weight: 500;
  }
 
  .file-selected {
    display: none;
    align-items: center;
    gap: 0.6rem;
    background: rgba(0,229,160,0.06);
    border: 1px solid var(--accent2);
    border-radius: 3px;
    padding: 0.6rem 0.9rem;
    font-size: 0.78rem;
    color: var(--accent);
    margin-bottom: 1.5rem;
  }
 
  .file-selected.visible { display: flex; }
 
  .file-selected .fname {
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
 
  .btn {
    width: 100%;
    background: var(--accent);
    color: #000;
    border: none;
    border-radius: 3px;
    padding: 0.9rem 1.5rem;
    font-family: var(--display);
    font-size: 0.9rem;
    font-weight: 800;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    cursor: pointer;
    transition: opacity 0.15s, transform 0.1s;
    position: relative;
    overflow: hidden;
  }
 
  .btn:hover:not(:disabled) { opacity: 0.88; }
  .btn:active:not(:disabled) { transform: scale(0.99); }
  .btn:disabled { opacity: 0.35; cursor: not-allowed; }
 
  .status {
    margin-top: 1.2rem;
    font-size: 0.75rem;
    letter-spacing: 0.04em;
    min-height: 1.2rem;
    display: flex;
    align-items: center;
    gap: 0.5rem;
  }
 
  .status.running { color: var(--muted); }
  .status.error   { color: var(--danger); }
  .status.done    { color: var(--accent); }
 
  .spinner {
    width: 12px; height: 12px;
    border: 1.5px solid var(--muted);
    border-top-color: var(--accent);
    border-radius: 50%;
    animation: spin 0.7s linear infinite;
    flex-shrink: 0;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
 
  .steps {
    margin-top: 2rem;
    display: flex;
    flex-direction: column;
    gap: 0.4rem;
  }
 
  .step {
    display: flex;
    align-items: center;
    gap: 0.7rem;
    font-size: 0.72rem;
    color: var(--muted);
    padding: 0.4rem 0;
    border-bottom: 1px solid var(--border);
    transition: color 0.3s;
  }
 
  .step:last-child { border-bottom: none; }
  .step.active { color: var(--text); }
  .step.done-step { color: var(--accent); }
 
  .step-dot {
    width: 6px; height: 6px;
    border-radius: 50%;
    background: var(--border);
    flex-shrink: 0;
    transition: background 0.3s;
  }
  .step.active .step-dot { background: var(--accent); box-shadow: 0 0 6px var(--accent); }
  .step.done-step .step-dot { background: var(--accent); }
 
  .footer {
    margin-top: 2rem;
    font-size: 0.65rem;
    color: var(--muted);
    letter-spacing: 0.06em;
    text-align: center;
  }
</style>
</head>
<body>
<div class="container">
  <div class="logo">ETP Fuel Distribution</div>
  <h1>FIFO<br><span>Engine</span></h1>
  <p class="subtitle">Upload the Investor Summary workbook.<br>The engine allocates invoices, costs, and inventory in one pass.</p>
 
  <div class="card">
    <div class="drop-zone" id="dropZone">
      <input type="file" id="fileInput" accept=".xlsx">
      <span class="drop-icon">⬆</span>
      <div class="drop-label">
        <strong>Click to upload</strong> or drag &amp; drop<br>
        .xlsx files only
      </div>
    </div>
 
    <div class="file-selected" id="fileSelected">
      <span>📄</span>
      <span class="fname" id="fileName"></span>
      <span style="color:var(--muted)">✓</span>
    </div>
 
    <button class="btn" id="runBtn" disabled onclick="runEngine()">Run FIFO Engine</button>
 
    <div class="status" id="status"></div>
 
    <div class="steps" id="steps" style="display:none">
      <div class="step" id="s1"><div class="step-dot"></div>Loading supplier invoice queues</div>
      <div class="step" id="s2"><div class="step-dot"></div>Allocating BOLs (Load Tracking order)</div>
      <div class="step" id="s3"><div class="step-dot"></div>Writing Purchase to BOL-RTB</div>
      <div class="step" id="s4"><div class="step-dot"></div>Running RTB → BTC FIFO</div>
      <div class="step" id="s5"><div class="step-dot"></div>Building FIFO sheet</div>
      <div class="step" id="s6"><div class="step-dot"></div>Updating Overall Summary</div>
    </div>
  </div>
 
  <div class="footer">ETP · Internal Tool · All data stays on-server</div>
</div>
 
<script>
  const fileInput  = document.getElementById('fileInput');
  const dropZone   = document.getElementById('dropZone');
  const fileSelected = document.getElementById('fileSelected');
  const fileName   = document.getElementById('fileName');
  const runBtn     = document.getElementById('runBtn');
  const statusEl   = document.getElementById('status');
  const stepsEl    = document.getElementById('steps');
  let selectedFile = null;
 
  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) setFile(fileInput.files[0]);
  });
 
  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault(); dropZone.classList.remove('dragover');
    if (e.dataTransfer.files[0]) setFile(e.dataTransfer.files[0]);
  });
 
  function setFile(f) {
    selectedFile = f;
    fileName.textContent = f.name;
    fileSelected.classList.add('visible');
    runBtn.disabled = false;
    statusEl.className = 'status';
    statusEl.innerHTML = '';
    stepsEl.style.display = 'none';
    document.querySelectorAll('.step').forEach(s => s.className = 'step');
  }
 
  function setStep(n) {
    for (let i = 1; i <= 6; i++) {
      const el = document.getElementById('s' + i);
      if (i < n)       el.className = 'step done-step';
      else if (i === n) el.className = 'step active';
      else             el.className = 'step';
    }
  }
 
  async function runEngine() {
    if (!selectedFile) return;
    runBtn.disabled = true;
    stepsEl.style.display = 'flex';
    statusEl.className = 'status running';
    statusEl.innerHTML = '<div class="spinner"></div> Processing…';
 
    const delays = [0, 400, 900, 1400, 2000, 2600];
    delays.forEach((d, i) => setTimeout(() => setStep(i + 1), d));
 
    const fd = new FormData();
    fd.append('file', selectedFile);
 
    try {
      const res = await fetch('/process', { method: 'POST', body: fd });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.error || 'Server error');
      }
      setStep(7);
      statusEl.className = 'status done';
      statusEl.innerHTML = '✓ Done — downloading…';
 
      const blob = await res.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href = url;
      a.download = selectedFile.name.replace('.xlsx', '_FIFO.xlsx');
      a.click();
      URL.revokeObjectURL(url);
    } catch(e) {
      statusEl.className = 'status error';
      statusEl.innerHTML = '✗ ' + e.message;
    } finally {
      runBtn.disabled = false;
    }
  }
</script>
</body>
</html>
"""
 
@app.route('/')
def index():
    return render_template_string(HTML)
 
@app.route('/process', methods=['POST'])
def process():
    from flask import Response
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
        except Exception as e:
            return jsonify(error=str(e)), 500
        with open(dst, 'rb') as fh:
            data = fh.read()
 
    return Response(
        data,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': 'attachment; filename=output.xlsx'}
    )
 
if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)
