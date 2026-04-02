@echo off
cd /d %~dp0
py -m pip install -r requirements.txt
py init_db.py
py csv_importer.py --csv sample_data\delegues_modele.csv
py app.py
