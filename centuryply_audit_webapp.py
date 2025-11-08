
import os, io, csv, uuid, traceback, math
from pathlib import Path
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify
import pandas as pd, numpy as np, matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, Frame, Image, Table, TableStyle

BASE_DIR = Path(__file__).parent.resolve()
UPLOAD_DIR = BASE_DIR / "uploads"
REPORTS_DIR = BASE_DIR / "reports"
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
for d in (UPLOAD_DIR, REPORTS_DIR, TEMPLATES_DIR, STATIC_DIR):
    d.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, template_folder=str(TEMPLATES_DIR), static_folder=str(STATIC_DIR))

ALLOWED_EXT = {'.xlsx', '.xls', '.xlsm'}
ACCENT = '#b41f23'  # CenturyPly red

def allowed_file(fn):
    return Path(fn).suffix.lower() in ALLOWED_EXT

@app.route('/')
def home():
    return render_template('centuryply_audit_with_scoring_final.html')

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

@app.route('/generate_report', methods=['GET'])
def generate_report():
    try:
        filename = request.args.get('filename')
        if filename:
            path = UPLOAD_DIR / filename
        else:
            files = sorted(UPLOAD_DIR.glob("*.xls*"), key=lambda p: p.stat().st_mtime, reverse=True)
            if not files:
                return "No upload found", 404
            path = files[0]
        if not path.exists():
            return "File not found", 404

        df = pd.read_excel(path, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

        # ensure percent column
        if ('%age' not in df.columns) and ('Total Score' in df.columns and 'Total' in df.columns):
            df['%age'] = df['Total Score'] / df['Total'] * 100

        # call duration normalization
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

        # team grouping
        team_col = None
        for cand in ['Franchise','Team Name','Team','franchise']:
            if cand in df.columns:
                team_col = cand; break
        name_col = None
        for cand in ['Name','Sl. Name','RM Name','Name ' , 'Name.']:
            if cand in df.columns:
                name_col = cand; break
        percent_col = '%age' if '%age' in df.columns else None

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        out_pdf = REPORTS_DIR / f"CenturyPly_Call_Audit_Report_{ts}.pdf"

        # create charts directory
        chart_imgs = []
        teams = [("All Teams", df)] if team_col is None else list(df.groupby(team_col))
        import matplotlib
        matplotlib.use('Agg')
        for tname, g in teams:
            # parameter averages
            params = ['Introduction','Project Registration','Product & Pricing requirement','Product FeedBack','Cross & upsell of product','Marketing Benefit','Redem tion','Call Closure','GTM Adherence','CRM Update','Softskill']
            avgs = {}
            for p in params:
                if p in g.columns:
                    avgs[p] = pd.to_numeric(g[p], errors='coerce').dropna().mean()
                else:
                    avgs[p] = None
            # draw bar chart
            labels = [k for k,v in avgs.items() if v is not None]
            vals = [v for v in avgs.values() if v is not None]
            if labels and len(vals)>0:
                fig, ax = plt.subplots(figsize=(8,2.2))
                ax.bar(labels, vals, color=ACCENT)
                ax.set_ylim(0,100)
                ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=8)
                plt.tight_layout()
                imgf = REPORTS_DIR / f"params_{str(tname)[:20].replace('/','_')}_{ts}.png"
                fig.savefig(imgf, dpi=150)
                plt.close(fig)
                chart_imgs.append(str(imgf))

            # top/bottom RM
            if name_col and percent_col and name_col in g.columns:
                rm_stats = g.groupby(name_col)[percent_col].agg(['mean','count']).rename(columns={'mean':'avg_pct'}).dropna().sort_values('avg_pct', ascending=False)
                top5 = rm_stats.head(5)
                bottom5 = rm_stats.tail(5)
                # chart top5
                if not top5.empty:
                    fig, ax = plt.subplots(figsize=(6,1.8))
                    top5['avg_pct'].plot(kind='bar', color=ACCENT, ax=ax)
                    ax.set_ylim(0,100); ax.set_xticklabels(top5.index, rotation=30, fontsize=8)
                    plt.tight_layout()
                    ptop = REPORTS_DIR / f"top5_{str(tname)[:20]}_{ts}.png"
                    fig.savefig(ptop, dpi=150); plt.close(fig); chart_imgs.append(str(ptop))
                if not bottom5.empty:
                    fig, ax = plt.subplots(figsize=(6,1.8))
                    bottom5['avg_pct'].plot(kind='bar', color='gray', ax=ax)
                    ax.set_ylim(0,100); ax.set_xticklabels(bottom5.index, rotation=30, fontsize=8)
                    plt.tight_layout()
                    pb = REPORTS_DIR / f"bottom5_{str(tname)[:20]}_{ts}.png"
                    fig.savefig(pb, dpi=150); plt.close(fig); chart_imgs.append(str(pb))

        # trend scatter
        trend_img = None
        if percent_col and df[percent_col].notna().sum()>5 and df['_call_min'].notna().sum()>5:
            sub = df[[percent_col,'_call_min']].dropna()
            fig, ax = plt.subplots(figsize=(6,3))
            ax.scatter(sub['_call_min'], sub[percent_col], alpha=0.6, color=ACCENT)
            try:
                m,b = np.polyfit(sub['_call_min'], sub[percent_col], 1)
                xs = np.linspace(sub['_call_min'].min(), sub['_call_min'].max(), 100)
                ax.plot(xs, m*xs+b, color='black', linewidth=1)
            except:
                pass
            ax.set_xlabel('Call duration (min)'); ax.set_ylabel('Audit %')
            plt.tight_layout()
            trend_img = REPORTS_DIR / f"trend_{ts}.png"
            fig.savefig(trend_img, dpi=150); plt.close(fig)
            chart_imgs.append(str(trend_img))

        # Build PDF with ReportLab (red-accent header, logo placeholder)
        c = canvas.Canvas(str(out_pdf), pagesize=A4)
        w,h = A4
        # header band
        c.setFillColor(colors.HexColor(ACCENT))
        c.rect(0, h-70, w, 70, stroke=0, fill=1)
        # logo placeholder top-left
        logo_path = STATIC_DIR / "logo.png"
        if logo_path.exists():
            try:
                c.drawImage(str(logo_path), 40, h-60, width=120, height=40, preserveAspectRatio=True, mask='auto')
            except:
                c.setFillColor(colors.white); c.setFont("Helvetica-Bold", 14); c.drawString(40, h-45, "CenturyPly")
        else:
            c.setFillColor(colors.white); c.setFont("Helvetica-Bold", 14); c.drawString(40, h-45, "CenturyPly (logo)")
        # title on header
        c.setFont("Helvetica-Bold", 16); c.setFillColor(colors.white)
        c.drawString(200, h-45, "Call Audit Analysis Report")
        c.setFont("Helvetica", 9); c.drawString(200, h-60, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        y = h-110
        styles = getSampleStyleSheet()

        # overview table
        if 'Total Score' in df.columns and percent_col:
            overall_avg = round(df[percent_col].dropna().mean(),2)
            overall_std = round(df[percent_col].dropna().std(),2)
        else:
            overall_avg = overall_std = 'N/A'
        c.setFillColor(colors.black); c.setFont("Helvetica-Bold", 12)
        c.drawString(40, y, f"Overall Summary: Avg % = {overall_avg}    Std Dev = {overall_std}")
        y -= 24

        # Insert parameter charts and team blocks
        for img in chart_imgs:
            if y < 140:
                c.showPage(); # redraw header on new page
                c.setFillColor(colors.HexColor(ACCENT)); c.rect(0, h-70, w, 70, stroke=0, fill=1)
                if logo_path.exists():
                    try: c.drawImage(str(logo_path), 40, h-60, width=120, height=40, preserveAspectRatio=True, mask='auto')
                    except: pass
                c.setFont("Helvetica-Bold", 16); c.setFillColor(colors.white); c.drawString(200, h-45, "Call Audit Analysis Report")
                y = h-110
            try:
                c.drawImage(ImageReader(img), 40, y-160, width=520, height=120, preserveAspectRatio=True)
                y -= 160
            except Exception:
                pass

        c.showPage()
        c.save()

        return send_file(str(out_pdf), as_attachment=True)
    except Exception as e:
        traceback.print_exc()
        return str(e), 500

@app.route('/reports')
def reports_list():
    files = sorted(Path('reports').glob('*.pdf'), key=lambda p:p.stat().st_mtime, reverse=True)
    return jsonify([p.name for p in files])

@app.route('/reports/<path:name>')
def download(name):
    p = REPORTS_DIR / name
    if not p.exists(): return "Not found", 404
    return send_file(str(p), as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
