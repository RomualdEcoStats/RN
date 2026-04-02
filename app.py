from pathlib import Path
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, abort, send_file, session, flash

from config import (
    FLASK_SECRET, ORG_NAME, ORG_WEBSITE, DEFAULT_STATUSES, ADMIN_USERNAME, ADMIN_PASSWORD,
    ORG_EMAIL, ORG_PHONE, PRESIDENT_NAME
)
from database import init_db, fetch_all, fetch_one, insert_mandate, update_mandate, execute
from security import generate_uid, generate_signature, verify_signature
from generator import generate_all

app = Flask(__name__)
app.secret_key = FLASK_SECRET
init_db()

def login_required(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if not session.get("admin_logged_in"):
            return redirect(url_for("login", next=request.path))
        return func(*args, **kwargs)
    return wrapper

@app.context_processor
def inject_globals():
    return {
        "ORG_NAME": ORG_NAME,
        "ORG_WEBSITE": ORG_WEBSITE,
        "ORG_EMAIL": ORG_EMAIL,
        "ORG_PHONE": ORG_PHONE,
        "PRESIDENT_NAME": PRESIDENT_NAME,
        "DEFAULT_STATUSES": DEFAULT_STATUSES,
    }

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            session["admin_logged_in"] = True
            flash("Connexion réussie.", "success")
            return redirect(request.args.get("next") or url_for("dashboard"))
        flash("Identifiants invalides.", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/")
@login_required
def dashboard():
    mandates = fetch_all("SELECT * FROM mandates ORDER BY id DESC")
    total = len(mandates)
    by_status = {s: 0 for s in DEFAULT_STATUSES}
    for m in mandates:
        if m["statut"] in by_status:
            by_status[m["statut"]] += 1
    return render_template("dashboard.html", mandates=mandates, total=total, by_status=by_status)

@app.route("/mandates/new", methods=["GET", "POST"])
@login_required
def create_mandate():
    if request.method == "POST":
        reference = request.form.get("reference","").strip()
        if not reference:
            prefix = request.form.get("reference_prefix", "RN").strip() or "RN"
            year = (request.form.get("date_emission","2026")[:4] or "2026")
            next_num = (fetch_one("SELECT COALESCE(MAX(id),0)+1 AS n FROM mandates")["n"])
            reference = f"{prefix}/{year}/ONG-RN/PCA/CSA/{int(next_num):04d}"

        payload = {
            "reference": reference,
            "mandate_uid": generate_uid(),
            "signature_token": "",
            "civilite": request.form.get("civilite",""),
            "nom": request.form.get("nom","").strip(),
            "prenom": request.form.get("prenom","").strip(),
            "adresse": request.form.get("adresse","").strip(),
            "telephone": request.form.get("telephone","").strip(),
            "email": request.form.get("email","").strip(),
            "fonction": request.form.get("fonction","Délégué").strip(),
            "zone_intervention": request.form.get("zone_intervention","").strip(),
            "date_emission": request.form.get("date_emission","").strip(),
            "date_expiration": request.form.get("date_expiration","").strip(),
            "statut": request.form.get("statut","actif").strip(),
            "ville_signature": request.form.get("ville_signature","Strasbourg").strip(),
            "pays": request.form.get("pays","Bénin").strip(),
            "preferences_affichage": request.form.get("preferences_affichage","").strip(),
            "notes": request.form.get("notes","").strip(),
        }
        payload["signature_token"] = generate_signature(payload["reference"], payload["mandate_uid"])
        qr, docx, pdf = generate_all(payload)
        payload["qr_path"] = str(qr)
        payload["docx_path"] = str(docx)
        payload["pdf_path"] = str(pdf)
        insert_mandate(payload)
        flash("Mandat créé et documents générés.", "success")
        return redirect(url_for("dashboard"))
    return render_template("form.html", mode="create", mandate=None)

@app.route("/mandates/<path:reference>/edit", methods=["GET", "POST"])
@login_required
def edit_mandate(reference):
    mandate = fetch_one("SELECT * FROM mandates WHERE reference=?", (reference,))
    if not mandate:
        abort(404)

    if request.method == "POST":
        payload = {
            "civilite": request.form.get("civilite",""),
            "nom": request.form.get("nom","").strip(),
            "prenom": request.form.get("prenom","").strip(),
            "adresse": request.form.get("adresse","").strip(),
            "telephone": request.form.get("telephone","").strip(),
            "email": request.form.get("email","").strip(),
            "fonction": request.form.get("fonction","Délégué").strip(),
            "zone_intervention": request.form.get("zone_intervention","").strip(),
            "date_emission": request.form.get("date_emission","").strip(),
            "date_expiration": request.form.get("date_expiration","").strip(),
            "statut": request.form.get("statut","actif").strip(),
            "ville_signature": request.form.get("ville_signature","Strasbourg").strip(),
            "pays": request.form.get("pays","Bénin").strip(),
            "preferences_affichage": request.form.get("preferences_affichage","").strip(),
            "notes": request.form.get("notes","").strip(),
        }
        current = dict(mandate)
        current.update(payload)
        qr, docx, pdf = generate_all(current)
        payload["qr_path"] = str(qr)
        payload["docx_path"] = str(docx)
        payload["pdf_path"] = str(pdf)
        update_mandate(reference, payload)
        flash("Mandat modifié et documents régénérés.", "success")
        return redirect(url_for("dashboard"))

    return render_template("form.html", mode="edit", mandate=mandate)

@app.route("/mandates/<path:reference>/status", methods=["POST"])
@login_required
def update_status(reference):
    status = request.form.get("statut","actif")
    if status not in DEFAULT_STATUSES:
        abort(400)
    mandate = fetch_one("SELECT * FROM mandates WHERE reference=?", (reference,))
    if not mandate:
        abort(404)
    current = dict(mandate)
    payload = {"statut": status}
    current["statut"] = status
    qr, docx, pdf = generate_all(current)
    payload["qr_path"] = str(qr)
    payload["docx_path"] = str(docx)
    payload["pdf_path"] = str(pdf)
    updated = update_mandate(reference, payload)
    if not updated:
        abort(404)
    flash("Statut mis à jour et document régénéré.", "success")
    return redirect(url_for("dashboard"))

@app.route("/mandates/<path:reference>/regen")
@login_required
def regen(reference):
    mandate = fetch_one("SELECT * FROM mandates WHERE reference=?", (reference,))
    if not mandate:
        abort(404)
    current = dict(mandate)
    qr, docx, pdf = generate_all(current)
    update_mandate(reference, {"qr_path": str(qr), "docx_path": str(docx), "pdf_path": str(pdf)})
    flash("Documents régénérés.", "success")
    return redirect(url_for("dashboard"))

@app.route("/mandates/<path:reference>/delete", methods=["POST"])
@login_required
def delete_mandate(reference):
    execute("DELETE FROM mandates WHERE reference=?", (reference,))
    flash("Mandat supprimé.", "success")
    return redirect(url_for("dashboard"))

@app.route("/verify")
def verify():
    reference = request.args.get("ref","")
    uid = request.args.get("uid","")
    sig = request.args.get("sig","")
    mandate = fetch_one("SELECT * FROM mandates WHERE reference=?", (reference,))
    valid = False
    if mandate:
        valid = uid == mandate["mandate_uid"] and verify_signature(reference, uid, sig)
    return render_template("verify.html", mandate=mandate, valid=valid, ref=reference)

@app.route("/documents/<kind>/<path:reference>")
@login_required
def open_document(kind, reference):
    mandate = fetch_one("SELECT * FROM mandates WHERE reference=?", (reference,))
    if not mandate:
        abort(404)
    if kind == "pdf":
        path = mandate["pdf_path"]
    elif kind == "docx":
        path = mandate["docx_path"]
    elif kind == "qr":
        path = mandate["qr_path"]
    else:
        abort(404)
    if not path or not Path(path).exists():
        abort(404)
    return send_file(path, as_attachment=False)

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)
