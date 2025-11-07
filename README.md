# 🏢 CenturyPly Corporate — QA Audit & Scoring System

Copyright © CenturyPly Corporate

Overview:
A Flask-based web platform for call-audit evaluation and QA performance tracking.

Routes:
- / : Scoring dashboard (default)
- /scoring : Scoring page
- /upload : Upload Excel audit files
- /reports : View generated reports
- /save_status : POST endpoint to save QA status logs

Local run:
1. pip install -r requirements.txt
2. python centuryply_audit_webapp.py
3. Open http://127.0.0.1:5000/

Deployment:
- Render: Push repo & use render.yaml
- Railway: Connect repo and deploy
- Deta: space push
