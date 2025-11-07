\
# centuryply_audit_webapp.py
import os, csv, uuid, io, datetime as dt
from pathlib import Path
from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify, send_from_directory
import pandas as pd

BASE_DIR = Path(__file__).parent.resolve()
UPLOAD_DIR = BASE_DIR / "uploads"
REPORTS_DIR = BASE_DIR / "reports"
STATIC_DIR = BASE_DIR / "static"
for d in (UPLOAD_DIR, REPORTS_DIR, STATIC_DIR):
    d.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"), static_folder=str(STATIC_DIR))
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

ALLOWED_EXT = {'.xlsx', '.xlsm', '.xls'}

def allowed_file(filename):
    return Path(filename).suffix.lower() in ALLOWED_EXT

# --- Routes ---
@app.route('/')
def home():
    return render_template('centuryply_audit_with_scoring_final.html')

@app.route('/scoring')
def scoring_page():
    return render_template('centuryply_audit_with_scoring_final.html')

@app.route('/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not f or f.filename == "":
        return "No file", 400
    if not allowed_file(f.filename):
        return "Invalid file", 400
    fn = f"{uuid.uuid4().hex}_{f.filename}"
    path = UPLOAD_DIR / fn
    f.save(str(path))
    # Minimal processing - store original for report generation pipeline
    return redirect(url_for('ready'))

@app.route('/ready')
def ready():
    return render_template('ready.html')

@app.route('/reports')
def reports():
    files = []
    for p in REPORTS_DIR.glob("*.pdf"):
        files.append({"name": p.name, "pdf_url": url_for('download_report', filename=p.name), "modified": dt.datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")})
    return render_template('reports.html', reports=files)

@app.route('/reports/<path:filename>')
def download_report(filename):
    p = REPORTS_DIR / filename
    if not p.exists():
        return "Not found", 404
    return send_file(str(p), as_attachment=True)

# --- Save status route with Call Date & Audit Date ---
@app.route('/save_status', methods=['POST'])
def save_status():
    try:
        data = request.get_json() or {}
        rm_name = data.get('rm_name', 'Unknown RM')
        status = data.get('status', 'N/A')
        evaluator = data.get('evaluator', 'Unknown QA')
        call_date = data.get('call_date', 'N/A')
        audit_date = data.get('audit_date', 'N/A')
        timestamp = dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        REPORTS_DIR.mkdir(parents=True, exist_ok=True)
        log_file = REPORTS_DIR / "qa_status_log.csv"
        header_needed = not log_file.exists()
        with open(log_file, 'a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if header_needed:
                writer.writerow(["Timestamp", "Evaluator", "RM Name", "Status", "Call Date", "Audit Date"])
            writer.writerow([timestamp, evaluator, rm_name, status, call_date, audit_date])
        return jsonify({"success": True, "message": f'Status \"{status}\" saved successfully for {rm_name}.'})
    except Exception as e:
        return jsonify({"success": False, "message": f"Error saving QA status: {str(e)}"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
