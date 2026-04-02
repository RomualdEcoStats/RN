# Pack V3 complet — Attestations de mandat ONG RN

## Contenu
- Back-office premium : création, édition, suppression, mise à jour de statut
- Régénération automatique du PDF, du DOCX et du QR code
- Vérification publique de l’authenticité
- Déploiement production prêt pour domaine réel
- Login administrateur simple (à remplacer immédiatement avant mise en ligne)

## Installation locale Windows
```powershell
py -m pip install -r requirements.txt
py init_db.py
py csv_importer.py --csv sample_data\delegues_modele.csv
py app.py
```
Puis ouvrir : http://127.0.0.1:5000

Identifiants par défaut :
- admin
- admin123

## Déploiement production
1. Décompresser sur serveur Linux
2. Configurer `.env` au besoin via variables d'environnement
3. Installer :
```bash
pip3 install -r requirements.txt
python3 init_db.py
gunicorn -w 2 -b 127.0.0.1:5000 app:app
```
4. Nginx reverse proxy vers `127.0.0.1:5000`
5. SSL avec Certbot
6. Remplacer `BASE_VERIFY_URL` par votre URL publique, ex. :
`https://www.renaitredenouveau.org/verify`

## Important
- Le QR code de test local ne fonctionnera pas sur téléphone tant que `BASE_VERIFY_URL` pointe vers `127.0.0.1`.
- Après chaque changement de statut ou édition, les documents sont régénérés.
