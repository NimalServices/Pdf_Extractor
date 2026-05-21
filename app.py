"""
Bank Statement Extractor — Flask Web App
Run: pip install flask camelot-py[cv] pandas openpyxl ghostscript
Then: python app.py
Open: http://localhost:5000
"""

from flask import Flask, request, send_file, jsonify
import camelot
import pandas as pd
import re
import os
import tempfile
import io

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB limit

# ── Extraction logic (unchanged) ─────────────────────────────────────────────

PAYMENT_TYPES = {'DD', 'CR', 'BP', 'OBP', 'VIS', 'TFR', 'CHQ', 'FP', 'SO'}

def is_transaction_table(df):
    if df.shape[1] < 5:
        return False
    col_text = ' '.join(df.iloc[:, 1].astype(str).tolist())
    return any(pt in col_text for pt in PAYMENT_TYPES) or 'BALANCE BROUGHT FORWARD' in df.to_string()

def clean_transactions(df):
    rows = []
    current_date = ''
    i = 0

    while i < len(df):
        row = df.iloc[i]
        cols = [str(c).strip() for c in row]

        if any(skip in ' '.join(cols) for skip in [
            'BUSINESS CURRENT ACCOUNT', 'Pay m e nt', 'Paid out',
            'Date', 'Sheet', 'Account Nam', 'Sort'
        ]):
            i += 1
            continue

        date_val = cols[0] if re.match(r'\d{2}\s\w+\s\d{2}', cols[0]) else ''
        if date_val:
            current_date = date_val

        ptype = cols[1] if cols[1] in PAYMENT_TYPES else ''
        desc  = cols[2].strip()

        if not ptype and not date_val and i > 0:
            if rows:
                prev = rows[-1]
                prev['Description'] = (prev['Description'] + ' ' + desc).strip()
                for field, col_idx in [('Paid Out', 3), ('Paid In', 4), ('Balance', 5)]:
                    val = cols[col_idx] if len(cols) > col_idx else ''
                    if val and val not in ('.', ''):
                        prev[field] = val
            i += 1
            continue

        if 'BALANCE BROUGHT FORWARD' in desc or 'BALANCE CARRIED FORWARD' in desc:
            balance = cols[5] if len(cols) > 5 else (cols[4] if len(cols) > 4 else '')
            rows.append({
                'Date': current_date, 'Type': '', 'Description': desc,
                'Paid Out': '', 'Paid In': '', 'Balance': balance
            })
            i += 1
            continue

        paid_out = cols[3] if len(cols) > 3 else ''
        paid_in  = cols[4] if len(cols) > 4 else ''
        balance  = cols[5] if len(cols) > 5 else ''

        rows.append({
            'Date': current_date,
            'Type': ptype,
            'Description': desc,
            'Paid Out': paid_out,
            'Paid In': paid_in,
            'Balance': balance
        })
        i += 1

    return rows

# ── Routes ────────────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Statement Extractor</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500&display=swap" rel="stylesheet">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg: #0e0f11;
    --surface: #16181c;
    --surface2: #1e2126;
    --border: #2a2d34;
    --accent: #00e5a0;
    --accent-dim: rgba(0,229,160,0.12);
    --accent-dim2: rgba(0,229,160,0.06);
    --text: #e8eaed;
    --muted: #7a7f8a;
    --danger: #ff5f5f;
    --danger-dim: rgba(255,95,95,0.1);
  }

  body {
    background: var(--bg);
    color: var(--text);
    font-family: 'DM Sans', sans-serif;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 3rem 1.5rem;
  }

  .wordmark {
    font-family: 'DM Mono', monospace;
    font-size: 11px;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 2.5rem;
    opacity: 0.85;
  }

  h1 {
    font-size: clamp(1.8rem, 4vw, 2.6rem);
    font-weight: 300;
    letter-spacing: -0.03em;
    text-align: center;
    margin-bottom: 0.5rem;
    line-height: 1.2;
  }

  h1 span { color: var(--accent); font-weight: 500; }

  .subtitle {
    color: var(--muted);
    font-size: 0.9rem;
    text-align: center;
    margin-bottom: 3rem;
    font-weight: 300;
  }

  .card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 2rem;
    width: 100%;
    max-width: 540px;
  }

  /* Drop zone */
  .drop-zone {
    border: 1.5px dashed var(--border);
    border-radius: 12px;
    padding: 2.5rem 1.5rem;
    text-align: center;
    cursor: pointer;
    transition: border-color 0.2s, background 0.2s;
    background: var(--accent-dim2);
    position: relative;
  }
  .drop-zone:hover, .drop-zone.drag-over {
    border-color: var(--accent);
    background: var(--accent-dim);
  }
  .drop-zone input[type=file] {
    position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%;
  }

  .drop-icon {
    width: 48px; height: 48px;
    margin: 0 auto 1rem;
    border-radius: 12px;
    background: var(--accent-dim);
    display: flex; align-items: center; justify-content: center;
  }
  .drop-icon svg { stroke: var(--accent); }

  .drop-label {
    font-size: 0.9rem; color: var(--muted);
    line-height: 1.6;
  }
  .drop-label strong { color: var(--text); font-weight: 500; }

  .file-pill {
    display: none;
    align-items: center;
    gap: 10px;
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 10px 14px;
    margin-top: 1rem;
    font-size: 0.85rem;
  }
  .file-pill.visible { display: flex; }
  .file-pill .name { flex: 1; color: var(--text); font-family: 'DM Mono', monospace; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .file-pill .size { color: var(--muted); font-family: 'DM Mono', monospace; }
  .file-pill .remove { background: none; border: none; color: var(--muted); cursor: pointer; padding: 0; line-height: 1; font-size: 16px; }
  .file-pill .remove:hover { color: var(--danger); }

  /* Format picker */
  .format-row {
    display: flex; gap: 10px; margin-top: 1.25rem;
  }
  .fmt-btn {
    flex: 1;
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 10px;
    font-family: 'DM Mono', monospace;
    font-size: 12px;
    color: var(--muted);
    cursor: pointer;
    letter-spacing: 0.08em;
    transition: border-color 0.15s, color 0.15s, background 0.15s;
  }
  .fmt-btn:hover { border-color: var(--accent); color: var(--text); }
  .fmt-btn.active { border-color: var(--accent); background: var(--accent-dim); color: var(--accent); }

  /* Extract button */
  .btn-extract {
    width: 100%;
    margin-top: 1.25rem;
    padding: 13px;
    background: var(--accent);
    color: #0a0f0a;
    font-family: 'DM Sans', sans-serif;
    font-size: 0.9rem;
    font-weight: 500;
    border: none;
    border-radius: 10px;
    cursor: pointer;
    letter-spacing: 0.02em;
    transition: opacity 0.15s, transform 0.1s;
    display: flex; align-items: center; justify-content: center; gap: 8px;
  }
  .btn-extract:hover { opacity: 0.9; }
  .btn-extract:active { transform: scale(0.99); }
  .btn-extract:disabled { opacity: 0.4; cursor: not-allowed; }

  /* Progress */
  .progress-wrap {
    display: none;
    margin-top: 1.25rem;
  }
  .progress-wrap.visible { display: block; }
  .progress-bar-track {
    height: 3px; background: var(--border); border-radius: 99px; overflow: hidden;
  }
  .progress-bar-fill {
    height: 100%; background: var(--accent); border-radius: 99px;
    width: 0%; transition: width 0.4s ease;
    animation: indeterminate 1.4s ease infinite;
  }
  @keyframes indeterminate {
    0%   { margin-left: -40%; width: 40%; }
    60%  { margin-left: 100%; width: 40%; }
    100% { margin-left: 100%; width: 40%; }
  }
  .progress-label {
    font-size: 12px; color: var(--muted); margin-top: 8px;
    font-family: 'DM Mono', monospace; letter-spacing: 0.04em;
  }

  /* Error / success */
  .msg {
    display: none; margin-top: 1rem;
    padding: 12px 14px; border-radius: 8px;
    font-size: 0.85rem; line-height: 1.5;
  }
  .msg.visible { display: block; }
  .msg.error { background: var(--danger-dim); border: 1px solid var(--danger); color: var(--danger); }
  .msg.success { background: var(--accent-dim); border: 1px solid var(--accent); color: var(--accent); }

  /* Stats row */
  .stats { display: none; gap: 1px; margin-top: 1.25rem; border-radius: 10px; overflow: hidden; border: 1px solid var(--border); }
  .stats.visible { display: flex; }
  .stat { flex: 1; background: var(--surface2); padding: 12px 14px; }
  .stat-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.12em; color: var(--muted); font-family: 'DM Mono', monospace; }
  .stat-val { font-size: 1.2rem; font-weight: 500; color: var(--text); margin-top: 2px; }

  /* Download */
  .btn-download {
    display: none; width: 100%; margin-top: 1rem;
    padding: 13px; border: 1.5px solid var(--accent);
    background: transparent; color: var(--accent);
    font-family: 'DM Sans', sans-serif; font-size: 0.9rem; font-weight: 500;
    border-radius: 10px; cursor: pointer; letter-spacing: 0.02em;
    transition: background 0.15s;
    align-items: center; justify-content: center; gap: 8px;
    text-decoration: none;
  }
  .btn-download.visible { display: flex; }
  .btn-download:hover { background: var(--accent-dim); }

  footer {
    margin-top: 3rem;
    font-size: 11px;
    color: var(--muted);
    font-family: 'DM Mono', monospace;
    letter-spacing: 0.06em;
    opacity: 0.5;
  }
</style>
</head>
<body>
<div class="wordmark">bank statement extractor</div>

<h1>Turn PDFs into<br><span>clean spreadsheets</span></h1>
<p class="subtitle">Upload an  business statement — get a structured Excel or CSV file instantly.</p>

<div class="card">
  <div class="drop-zone" id="dropZone">
    <input type="file" id="fileInput" accept=".pdf">
    <div class="drop-icon">
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
        <polyline points="14 2 14 8 20 8"/>
        <line x1="12" y1="18" x2="12" y2="12"/>
        <polyline points="9 15 12 12 15 15"/>
      </svg>
    </div>
    <p class="drop-label"><strong>Drop your PDF here</strong><br>or click to browse</p>
  </div>

  <div class="file-pill" id="filePill">
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;color:var(--accent)"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
    <span class="name" id="fileName"></span>
    <span class="size" id="fileSize"></span>
    <button class="remove" id="removeFile" title="Remove">✕</button>
  </div>

  <div class="format-row">
    <button class="fmt-btn active" data-fmt="xlsx">XLSX</button>
    <button class="fmt-btn" data-fmt="csv">CSV</button>
  </div>

  <button class="btn-extract" id="extractBtn" disabled>
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>
    Extract Transactions
  </button>

  <div class="progress-wrap" id="progressWrap">
    <div class="progress-bar-track"><div class="progress-bar-fill" id="progressFill"></div></div>
    <div class="progress-label" id="progressLabel">Reading PDF tables…</div>
  </div>

  <div class="msg" id="msgBox"></div>

  # <div class="stats" id="stats">
  #   <div class="stat"><div class="stat-label">Rows</div><div class="stat-val" id="statRows">—</div></div>
  #   <div class="stat"><div class="stat-label">Paid Out</div><div class="stat-val" id="statOut">—</div></div>
  #   <div class="stat"><div class="stat-label">Paid In</div><div class="stat-val" id="statIn">—</div></div>
  # </div>

  <a class="btn-download" id="downloadBtn" href="#">
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
    Download <span id="downloadLabel">XLSX</span>
  </a>
</div>

<footer>made by Nimal &nbsp;·&nbsp; <a href="tel:+94765809268" style="color:inherit;text-decoration:none;">+94 XX XXX XXXX</a></footer>

<script>
let selectedFile = null;
let selectedFmt  = 'xlsx';
let downloadUrl  = null;

const dropZone   = document.getElementById('dropZone');
const fileInput  = document.getElementById('fileInput');
const filePill   = document.getElementById('filePill');
const fileName   = document.getElementById('fileName');
const fileSize   = document.getElementById('fileSize');
const removeFile = document.getElementById('removeFile');
const extractBtn = document.getElementById('extractBtn');
const progressWrap = document.getElementById('progressWrap');
const progressLabel = document.getElementById('progressLabel');
const msgBox     = document.getElementById('msgBox');
const stats      = document.getElementById('stats');
const downloadBtn = document.getElementById('downloadBtn');
const downloadLabel = document.getElementById('downloadLabel');

function fmtSize(b) {
  if (b < 1024) return b + ' B';
  if (b < 1024*1024) return (b/1024).toFixed(1) + ' KB';
  return (b/1024/1024).toFixed(1) + ' MB';
}

function setFile(f) {
  selectedFile = f;
  fileName.textContent = f.name;
  fileSize.textContent = fmtSize(f.size);
  filePill.classList.add('visible');
  extractBtn.disabled = false;
  clearResult();
}

function clearResult() {
  msgBox.className = 'msg'; msgBox.textContent = '';
  stats.classList.remove('visible');
  downloadBtn.classList.remove('visible');
  progressWrap.classList.remove('visible');
  if (downloadUrl) { URL.revokeObjectURL(downloadUrl); downloadUrl = null; }
}

removeFile.addEventListener('click', () => {
  selectedFile = null;
  fileInput.value = '';
  filePill.classList.remove('visible');
  extractBtn.disabled = true;
  clearResult();
});

fileInput.addEventListener('change', () => { if (fileInput.files[0]) setFile(fileInput.files[0]); });

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('drag-over');
  if (e.dataTransfer.files[0]?.type === 'application/pdf') setFile(e.dataTransfer.files[0]);
});

document.querySelectorAll('.fmt-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.fmt-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    selectedFmt = btn.dataset.fmt;
    downloadLabel.textContent = selectedFmt.toUpperCase();
    clearResult();
  });
});

const steps = ['Reading PDF tables…', 'Filtering transaction rows…', 'Merging split descriptions…', 'Building spreadsheet…'];
let stepIdx = 0, stepTimer;

function startProgress() {
  progressWrap.classList.add('visible');
  stepIdx = 0;
  progressLabel.textContent = steps[0];
  stepTimer = setInterval(() => {
    stepIdx = (stepIdx + 1) % steps.length;
    progressLabel.textContent = steps[stepIdx];
  }, 1800);
}

function stopProgress() {
  clearInterval(stepTimer);
  progressWrap.classList.remove('visible');
}

extractBtn.addEventListener('click', async () => {
  if (!selectedFile) return;
  clearResult();
  extractBtn.disabled = true;
  startProgress();

  const fd = new FormData();
  fd.append('file', selectedFile);
  fd.append('fmt', selectedFmt);

  try {
    const res = await fetch('/extract', { method: 'POST', body: fd });
    stopProgress();

    if (!res.ok) {
      const err = await res.json();
      msgBox.textContent = '⚠ ' + (err.error || 'Extraction failed.');
      msgBox.className = 'msg error visible';
    } else {
      const meta = JSON.parse(res.headers.get('X-Meta') || '{}');
      const blob = await res.blob();
      downloadUrl = URL.createObjectURL(blob);
      downloadBtn.href = downloadUrl;
      downloadBtn.download = selectedFile.name.replace('.pdf', '') + '_transactions.' + selectedFmt;
      downloadLabel.textContent = selectedFmt.toUpperCase();
      downloadBtn.classList.add('visible');

      document.getElementById('statRows').textContent = meta.rows ?? '—';
      document.getElementById('statOut').textContent  = meta.paid_out ? '£' + meta.paid_out : '—';
      document.getElementById('statIn').textContent   = meta.paid_in  ? '£' + meta.paid_in  : '—';
      stats.classList.add('visible');

      msgBox.textContent = '✓ Extracted ' + (meta.rows ?? '?') + ' transaction rows successfully.';
      msgBox.className = 'msg success visible';
    }
  } catch (e) {
    stopProgress();
    msgBox.textContent = '⚠ Network error: ' + e.message;
    msgBox.className = 'msg error visible';
  } finally {
    extractBtn.disabled = false;
  }
});
</script>
</body>
</html>
"""

@app.route('/')
def index():
    return HTML

@app.route('/extract', methods=['POST'])
def extract():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded.'}), 400

    f   = request.files['file']
    fmt = request.form.get('fmt', 'xlsx')

    if not f.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Only PDF files are supported.'}), 400

    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
        f.save(tmp.name)
        tmp_path = tmp.name

    try:
        stream_tables = camelot.read_pdf(tmp_path, pages='all', flavor='stream', suppress_stdout=True)

        all_rows = []
        for table in stream_tables:
            df = table.df
            if is_transaction_table(df):
                all_rows.extend(clean_transactions(df))

        if not all_rows:
            return jsonify({'error': 'No transaction rows found. Is this an  business statement?'}), 422

        result = pd.DataFrame(all_rows, columns=['Date', 'Type', 'Description', 'Paid Out', 'Paid In', 'Balance'])
        result['Date'] = result['Date'].replace('', pd.NA).ffill()

        # Compute summary stats (strip commas/£, ignore non-numeric)
        def col_sum(col):
            total = 0.0
            for v in result[col]:
                try:
                    total += float(str(v).replace(',', '').replace('£', '').strip())
                except ValueError:
                    pass
            return round(total, 2)

        meta = {
            'rows':     len(result),
            'paid_out': col_sum('Paid Out'),
            'paid_in':  col_sum('Paid In'),
        }

        buf = io.BytesIO()

        if fmt == 'csv':
            result.to_csv(buf, index=False)
            buf.seek(0)
            response = send_file(
                buf,
                mimetype='text/csv',
                as_attachment=True,
                download_name='transactions.csv'
            )
        else:
            with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                result.to_excel(writer, index=False, sheet_name='Transactions')
                ws = writer.sheets['Transactions']
                # Auto-size columns
                for col_cells in ws.columns:
                    length = max(len(str(c.value or '')) for c in col_cells)
                    ws.column_dimensions[col_cells[0].column_letter].width = min(length + 4, 50)
            buf.seek(0)
            response = send_file(
                buf,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='transactions.xlsx'
            )

        import json
        response.headers['X-Meta'] = json.dumps(meta)
        return response

    except Exception as e:
        return jsonify({'error': str(e)}), 500

    finally:
        os.unlink(tmp_path)


if __name__ == '__main__':
    print("\n  Bank Statement Extractor")
    print("  → http://localhost:5000\n")
    port = int(os.environ.get('PORT', 8080))
    app.run(debug=False, host='0.0.0.0', port=port)