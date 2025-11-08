
# CenturyPly Audit Reports - Web App (v4.0)

This repository contains the CenturyPly QA Audit web application.
It allows uploading audit Excel sheets and generates red-accent PDF reports,
provides a scoring UI (collapsible scoring framework) and saves QA observations.

## Quick local setup
```bash
unzip centuryply_corporate_final_build_v4.0_full.zip -d centuryply_app
cd centuryply_app
pip install -r requirements.txt
python centuryply_audit_webapp.py
# open http://127.0.0.1:5000
```

## Routes
- `/` — Scoring UI (homepage)
- `/scoring` — embedded scoring form (accordion)
- `/upload` — upload .xlsx audit file (generates report)
- `/generate_report?mode=full|team|rm` — generate PDF report
- `/save_status` — POST endpoint to save QA scoring (appends CSV)
- `/run_legacy` — POST endpoint to run legacy script located in /legacy

## Deploy on Render (Free)
1. Create a new Web Service on Render connected to this repository.
2. Build command: `pip install -r requirements.txt`
3. Start command: `gunicorn centuryply_audit_webapp:app`
4. Ensure `render.yaml` is present (already configured for free plan).

## Add your logo
Place your logo at `static/logo.png`. The app will show it in the top-left.

---
