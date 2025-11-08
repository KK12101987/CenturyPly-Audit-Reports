
import os, io, csv, uuid, traceback, math
from pathlib import Path
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify
import pandas as pd, numpy as np, matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

BASE_DIR = Path(__file__).parent.resolve()
UPLOAD_DIR = BASE_DIR / "uploads"
REPORTS_DIR = BASE_DIR / "reports"
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
for d in (UPLOAD_DIR, REPORTS_DIR, TEMPLATES_DIR, STATIC_DIR):
    d.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, template_folder=str(TEMPLATES_DIR), static_folder=str(STATIC_DIR))

ALLOWED_EXT = {'.xlsx', '.xls', '.xlsm'}
ACCENT = '#b41f23'

def allowed_file(fn):
    return Path(fn).suffix.lower() in ALLOWED_EXT

@app.route('/')
def home():
    return render_template('centuryply_audit_with_scoring_final.html')

@app.route('/scoring')
def scoring():
    return render_template('scoring_form_embed.html')

@app.route('/upload', methods=['GET','POST'])
def upload():
    if request.method == 'GET':
        return render_template('upload.html')
    f = request.files.get('file')
    if not f or f.filename == "":
        return "No file", 400
    if not allowed_file(f.filename):
        return "Unsupported", 400
    dest = UPLOAD_DIR / f"{uuid.uuid4().hex}_{f.filename}"
    f.save(dest)
    return redirect(url_for('generate_report', filename=dest.name))

@app.route('/save_status', methods=['POST'])
def save_status():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"success": False, "message": "No JSON received"}), 400
        REPORTS_DIR.mkdir(parents=True, exist_ok=True)
        csv_path = REPORTS_DIR / "qa_status_log.csv"
        headers = ["timestamp","evaluator","rm_name","team_name","rm_captain","date_of_audit","aid_mobile","call_date",
                   "Introduction","Project Registration","Product & Pricing requirement","Product FeedBack","Cross & upsell of product",
                   "Marketing Benefit","Redem tion","Call Closure","GTM Adherence","CRM Update","Softskill","Total Score","Total","%age","Audit Observation","status"]
        write_header = not csv_path.exists()
        with open(csv_path, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if write_header:
                writer.writerow(headers)
            row = [datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                   data.get("evaluator",""),
                   data.get("rm_name",""),
                   data.get("team_name",""),
                   data.get("rm_captain",""),
                   data.get("date_of_audit",""),
                   data.get("aid_mobile",""),
                   data.get("call_date",""),
                   data.get("Introduction",""),
                   data.get("Project Registration",""),
                   data.get("Product & Pricing requirement",""),
                   data.get("Product FeedBack",""),
                   data.get("Cross & upsell of product",""),
                   data.get("Marketing Benefit",""),
                   data.get("Redem tion",""),
                   data.get("Call Closure",""),
                   data.get("GTM Adherence",""),
                   data.get("CRM Update",""),
                   data.get("Softskill",""),
                   data.get("Total Score",""),
                   data.get("Total",""),
                   data.get("%age",""),
                   data.get("Audit Observation",""),
                   data.get("status","")]
            writer.writerow(row)
        return jsonify({"success": True, "message": "Saved successfully"})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"success": False, "message": str(e)}), 500

@app.route('/run_legacy', methods=['POST'])
def run_legacy():
    # Placeholder: run a legacy script if exists in /legacy, return its stdout
    legacy_dir = BASE_DIR / "legacy"
    out_lines = []
    try:
        script = None
        for p in legacy_dir.glob("*.py"):
            script = p
            break
        if script:
            import subprocess, sys, shlex, tempfile
            proc = subprocess.run([sys.executable, str(script)], capture_output=True, text=True, cwd=str(legacy_dir), timeout=60)
            out = proc.stdout + "\\n" + proc.stderr
            return jsonify({"success": True, "output": out})
        else:
            return jsonify({"success": False, "message": "No legacy script found in /legacy"})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"success": False, "message": str(e)})

@app.route('/generate_report', methods=['GET'])
def generate_report():
    from flask import request
    mode = request.args.get("mode", "full").lower()
    filename = request.args.get("filename")
    # select file
    if filename:
        path = UPLOAD_DIR / filename
        if not path.exists():
            return "File not found", 404
    else:
        files = sorted(UPLOAD_DIR.glob("*.xls*"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not files:
            return "No upload found", 404
        path = files[0]
    # read dataframe
    try:
        df = pd.read_excel(path, engine='openpyxl')
    except Exception as e:
        return f"Failed to read Excel: {e}", 500
    df.columns = [str(c).strip() for c in df.columns]
    # ensure %age exists
    if ('%age' not in df.columns) and ('Total Score' in df.columns and 'Total' in df.columns):
        df['%age'] = df['Total Score'] / df['Total'] * 100
    # normalize call duration
    dur_col = None
    for cand in ['Call Duration','CallDuration','Duration','call duration']:
        if cand in df.columns:
            dur_col = cand; break
    if dur_col:
        def parse(x):
            try:
                if pd.isna(x): return np.nan
                s = str(x)
                if ':' in s:
                    parts = [float(p) for p in s.split(':')]
                    if len(parts)==3:
                        return parts[0]*60 + parts[1] + parts[2]/60.0
                    elif len(parts)==2:
                        return parts[0] + parts[1]/60.0
                return float(s)
            except:
                return np.nan
        df['_call_min'] = df[dur_col].apply(parse)
    else:
        df['_call_min'] = np.nan

    # select mode
    if mode == 'full':
        out_pdf = REPORTS_DIR / f"CenturyPly_Full_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        _generate_full_pdf(df, out_pdf)
        return send_file(str(out_pdf), as_attachment=True)
    elif mode == 'team':
        out_pdf = REPORTS_DIR / f"CenturyPly_Teamwise_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        _generate_teamwise_pdf(df, out_pdf)
        return send_file(str(out_pdf), as_attachment=True)
    elif mode == 'rm':
        out_pdf = REPORTS_DIR / f"CenturyPly_RM_Summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        _generate_rm_pdf(df, out_pdf)
        return send_file(str(out_pdf), as_attachment=True)
    else:
        return "Invalid mode", 400

# ---------------- PDF helpers ----------------
def _draw_header(c, title):
    w,h = A4
    c.setFillColor(ACCENT)
    c.rect(0, h-70, w, 70, stroke=0, fill=1)
    logo = STATIC_DIR / "logo.png"
    c.setFillColorRGB(1,1,1)
    try:
        if logo.exists():
            c.drawImage(str(logo), 40, h-60, width=120, height=40, preserveAspectRatio=True, mask='auto')
        else:
            c.setFont("Helvetica-Bold", 14); c.drawString(40, h-45, "CenturyPly")
    except Exception:
        c.setFont("Helvetica-Bold", 14); c.drawString(40, h-45, "CenturyPly")
    c.setFont("Helvetica-Bold", 16); c.drawString(200, h-45, title)
    c.setFont("Helvetica", 9); c.drawString(200, h-60, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

def _generate_full_pdf(df, out_pdf):
    c = canvas.Canvas(str(out_pdf), pagesize=A4)
    w,h = A4
    _draw_header(c, "Call Audit Analysis Report - Full")
    y = h-100
    # overall summary
    percent_col = '%age' if '%age' in df.columns else None
    overall_avg = round(df[percent_col].dropna().mean(),2) if percent_col else 'N/A'
    c.setFont("Helvetica-Bold", 12); c.setFillColorRGB(0,0,0); c.drawString(40, y, f"Overall Average %: {overall_avg}")
    y -= 24
    # parameter averages (if present)
    params = ['Introduction','Project Registration','Product & Pricing requirement','Product FeedBack','Cross & upsell of product','Marketing Benefit','Redem tion','Call Closure','GTM Adherence','CRM Update','Softskill']
    param_avgs = {}
    for p in params:
        if p in df.columns:
            try:
                param_avgs[p] = round(pd.to_numeric(df[p], errors='coerce').dropna().mean(),2)
            except:
                param_avgs[p] = None
        else:
            param_avgs[p] = None
    # draw param chart image
    labels = [k for k,v in param_avgs.items() if v is not None]
    vals = [v for v in param_avgs.values() if v is not None]
    if labels and any([v is not None for v in vals]):
        import matplotlib
        matplotlib.use('Agg')
        fig, ax = plt.subplots(figsize=(8,2.2))
        ax.bar(labels, vals, color=ACCENT)
        ax.set_ylim(0,100)
        ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=8)
        plt.tight_layout()
        imgp = REPORTS_DIR / f"params_{uuid.uuid4().hex}.png"
        fig.savefig(imgp, dpi=150); plt.close(fig)
        try:
            c.drawImage(ImageReader(str(imgp)), 40, y-160, width=520, height=120, preserveAspectRatio=True)
            y -= 170
        except:
            pass

    # team-wise summaries
    team_col = None
    for cand in ['Franchise','Team Name','Team','franchise']:
        if cand in df.columns:
            team_col = cand; break
    if team_col:
        groups = df.groupby(team_col)
        for tname, g in groups:
            if y < 160:
                c.showPage(); _draw_header(c, "Call Audit Analysis Report - Full"); y = h-100
            c.setFont("Helvetica-Bold", 11); c.drawString(40, y, f"Team: {tname} - Audits: {len(g)}")
            y -= 16
            pct_avg = round(g['%age'].dropna().mean(),2) if '%age' in g.columns else 'N/A'
            c.setFont("Helvetica", 10); c.drawString(50, y, f"Avg %: {pct_avg}  Avg Call Duration (min): {round(g['_call_min'].dropna().mean(),2) if '_call_min' in g.columns else 'N/A'}")
            y -= 18
    c.showPage(); c.save()

def _generate_teamwise_pdf(df, out_pdf):
    c = canvas.Canvas(str(out_pdf), pagesize=A4)
    _draw_header(c, "Call Audit - Team-wise Report")
    w,h = A4; y = h-100
    team_col = None
    for cand in ['Franchise','Team Name','Team','franchise']:
        if cand in df.columns:
            team_col = cand; break
    if not team_col:
        c.setFont("Helvetica", 12); c.drawString(40, y, "No team column found in data"); c.showPage(); c.save(); return
    agg = df.groupby(team_col)['%age'].agg(['mean','median','std','count']).reset_index()
    x = 40; row_h = 16
    c.setFont("Helvetica-Bold", 10); c.drawString(x, y, "Team"); c.drawString(200, y, "Avg %"); c.drawString(300,y,"Median"); c.drawString(380,y,"Std"); c.drawString(450,y,"Count"); y -= row_h
    c.setFont("Helvetica", 9)
    for _,row in agg.iterrows():
        c.drawString(x, y, str(row[team_col])); c.drawString(200, y, str(round(row['mean'],2))); c.drawString(300,y,str(round(row['median'],2))); c.drawString(380,y,str(round(row['std'],2))); c.drawString(450,y,str(int(row['count']))); y -= row_h
        if y < 80:
            c.showPage(); _draw_header(c, "Call Audit - Team-wise Report"); y = h-100
    c.showPage(); c.save()

def _generate_rm_pdf(df, out_pdf):
    c = canvas.Canvas(str(out_pdf), pagesize=A4)
    _draw_header(c, "Call Audit - RM Summary")
    w,h = A4; y = h-100
    name_col = None
    for cand in ['Name','Sl. Name','RM Name','Name ' , 'Name.']:
        if cand in df.columns:
            name_col = cand; break
    if not name_col:
        c.setFont("Helvetica", 12); c.drawString(40, y, "No RM name column found"); c.showPage(); c.save(); return
    subset = df[[name_col,'%age']].dropna()
    agg = subset.groupby(name_col)['%age'].agg(['mean','count']).reset_index().sort_values('mean', ascending=False)
    x = 40; row_h = 16
    c.setFont("Helvetica-Bold", 10); c.drawString(x, y, "RM"); c.drawString(300,y,"Avg %"); c.drawString(380,y,"Audits"); y -= row_h
    c.setFont("Helvetica", 9)
    for _,row in agg.iterrows():
        c.drawString(x, y, str(row[name_col])); c.drawString(300, y, str(round(row['mean'],2))); c.drawString(380,y,str(int(row['count']))); y -= row_h
        if y < 80:
            c.showPage(); _draw_header(c, "Call Audit - RM Summary"); y = h-100
    c.showPage(); c.save()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
