from pathlib import Path
from functools import wraps
from datetime import timedelta
import os

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    abort,
    send_file,
    session,
    flash,
)
from werkzeug.middleware.proxy_fix import ProxyFix

from config import (
    FLASK_SECRET,
    ORG_NAME,
    ORG_WEBSITE,
    DEFAULT_STATUSES,
    ADMIN_USERNAME,
    ADMIN_PASSWORD,
    ORG_EMAIL,
    ORG_PHONE,
    PRESIDENT_NAME,
)
from database import init_db, fetch_all, fetch_one, insert_mandate, update_mandate, execute
from security import generate_uid, generate_signature, verify_signature
from generator import generate_all


app = Flask(__name__)
app.secret_key = FLASK_SECRET

# Important derrière Render / proxy HTTPS
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=1)

# Détection environnement Render
IS_RENDER = (
    os.getenv("RENDER", "").lower() == "true"
    or os.getenv("RENDER_EXTERNAL_URL") is not None
)

# Configuration session / cookies
app.config.update(
    SECRET_KEY=FLASK_SECRET,
    SESSION_COOKIE_NAME="ong_rn_session",
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=IS_RENDER,
    PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
    PREFERRED_URL_SCHEME="https" if IS_RENDER else "http",
    TEMPLATES_AUTO_RELOAD=False,
)

init_db()


def is_safe_next_url(target: str) -> bool:
    """
    Empêche les redirections externes malveillantes.
    """
    if not target:
        return False
    return target.startswith("/")


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


@app.route("/health")
def health():
    return {"status": "ok"}, 200


@app.route("/favicon.ico")
def favicon():
    return abort(204)


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        next_url = request.args.get("next", "")

        if (
            username.strip() == str(ADMIN_USERNAME).strip()
            and password.strip() == str(ADMIN_PASSWORD).strip()
        ):
            session.clear()
            session["admin_logged_in"] = True
            session["admin_username"] = username.strip()
            session.permanent = True

            flash("Connexion réussie.", "success")

            if is_safe_next_url(next_url):
                return redirect(next_url)
            return redirect(url_for("dashboard"))

        flash("Identifiants invalides.", "danger")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("Déconnexion effectuée.", "success")
    return redirect(url_for("login"))


@app.route("/")
@login_required
def dashboard():
    mandates = fetch_all("SELECT * FROM mandates ORDER BY id DESC")
    total = len(mandates)
    by_status = {status: 0 for status in DEFAULT_STATUSES}

    for mandate in mandates:
        if mandate["statut"] in by_status:
            by_status[mandate["statut"]] += 1

    return render_template(
        "dashboard.html",
        mandates=mandates,
        total=total,
        by_status=by_status,
    )


@app.route("/mandates/new", methods=["GET", "POST"])
@login_required
def create_mandate():
    if request.method == "POST":
        reference = request.form.get("reference", "").strip()

        if not reference:
            prefix = request.form.get("reference_prefix", "RN").strip() or "RN"
            year = (request.form.get("date_emission", "2026")[:4] or "2026")
            next_num_row = fetch_one("SELECT COALESCE(MAX(id), 0) + 1 AS n FROM mandates")
            next_num = next_num_row["n"] if next_num_row else 1
            reference = f"{prefix}/{year}/ONG-RN/PCA/CSA/{int(next_num):04d}"

        payload = {
            "reference": reference,
            "mandate_uid": generate_uid(),
            "signature_token": "",
            "civilite": request.form.get("civilite", "").strip(),
            "nom": request.form.get("nom", "").strip(),
            "prenom": request.form.get("prenom", "").strip(),
            "adresse": request.form.get("adresse", "").strip(),
            "telephone": request.form.get("telephone", "").strip(),
            "email": request.form.get("email", "").strip(),
            "fonction": request.form.get("fonction", "Délégué").strip(),
            "zone_intervention": request.form.get("zone_intervention", "").strip(),
            "date_emission": request.form.get("date_emission", "").strip(),
            "date_expiration": request.form.get("date_expiration", "").strip(),
            "statut": request.form.get("statut", "actif").strip(),
            "ville_signature": request.form.get("ville_signature", "Strasbourg").strip(),
            "pays": request.form.get("pays", "Bénin").strip(),
            "preferences_affichage": request.form.get("preferences_affichage", "").strip(),
            "notes": request.form.get("notes", "").strip(),
        }

        payload["signature_token"] = generate_signature(
            payload["reference"],
            payload["mandate_uid"],
        )

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
    mandate = fetch_one("SELECT * FROM mandates WHERE reference = ?", (reference,))
    if not mandate:
        abort(404)

    if request.method == "POST":
        payload = {
            "civilite": request.form.get("civilite", "").strip(),
            "nom": request.form.get("nom", "").strip(),
            "prenom": request.form.get("prenom", "").strip(),
            "adresse": request.form.get("adresse", "").strip(),
            "telephone": request.form.get("telephone", "").strip(),
            "email": request.form.get("email", "").strip(),
            "fonction": request.form.get("fonction", "Délégué").strip(),
            "zone_intervention": request.form.get("zone_intervention", "").strip(),
            "date_emission": request.form.get("date_emission", "").strip(),
            "date_expiration": request.form.get("date_expiration", "").strip(),
            "statut": request.form.get("statut", "actif").strip(),
            "ville_signature": request.form.get("ville_signature", "Strasbourg").strip(),
            "pays": request.form.get("pays", "Bénin").strip(),
            "preferences_affichage": request.form.get("preferences_affichage", "").strip(),
            "notes": request.form.get("notes", "").strip(),
        }

        current = dict(mandate)
        current.update(payload)

        qr, docx, pdf = generate_all(current)
        payload["qr_path"] = str(qr)
        payload["docx_path"] = str(docx)
        payload["pdf_path"] = str(pdf)

        updated = update_mandate(reference, payload)
        if not updated:
            abort(404)

        flash("Mandat modifié et documents régénérés.", "success")
        return redirect(url_for("dashboard"))

    return render_template("form.html", mode="edit", mandate=mandate)


@app.route("/mandates/<path:reference>/status", methods=["POST"])
@login_required
def update_status(reference):
    status = request.form.get("statut", "actif").strip()

    if status not in DEFAULT_STATUSES:
        abort(400)

    mandate = fetch_one("SELECT * FROM mandates WHERE reference = ?", (reference,))
    if not mandate:
        abort(404)

    current = dict(mandate)
    current["statut"] = status

    qr, docx, pdf = generate_all(current)
    payload = {
        "statut": status,
        "qr_path": str(qr),
        "docx_path": str(docx),
        "pdf_path": str(pdf),
    }

    updated = update_mandate(reference, payload)
    if not updated:
        abort(404)

    flash("Statut mis à jour et document régénéré.", "success")
    return redirect(url_for("dashboard"))


@app.route("/mandates/<path:reference>/regen")
@login_required
def regen(reference):
    mandate = fetch_one("SELECT * FROM mandates WHERE reference = ?", (reference,))
    if not mandate:
        abort(404)

    current = dict(mandate)
    qr, docx, pdf = generate_all(current)

    updated = update_mandate(
        reference,
        {
            "qr_path": str(qr),
            "docx_path": str(docx),
            "pdf_path": str(pdf),
        },
    )
    if not updated:
        abort(404)

    flash("Documents régénérés.", "success")
    return redirect(url_for("dashboard"))


@app.route("/mandates/<path:reference>/delete", methods=["POST"])
@login_required
def delete_mandate(reference):
    deleted = execute("DELETE FROM mandates WHERE reference = ?", (reference,))
    if not deleted:
        abort(404)

    flash("Mandat supprimé.", "success")
    return redirect(url_for("dashboard"))


@app.route("/verify")
def verify():
    from urllib.parse import unquote

    reference = request.args.get("ref", "")
    uid = request.args.get("uid", "")
    sig = request.args.get("sig", "")

    reference = unquote(reference).strip()
    uid = unquote(uid).strip()
    sig = unquote(sig).strip()

    print("VERIFY_REF_RAW:", repr(reference))
    print("VERIFY_UID_RAW:", repr(uid))
    print("VERIFY_SIG_RAW:", repr(sig))

    mandate = fetch_one(
        "SELECT * FROM mandates WHERE TRIM(reference) = TRIM(?)",
        (reference,)
    )

    all_refs = fetch_all("SELECT reference FROM mandates ORDER BY id DESC")
    print("DB_REFERENCES:", [row["reference"] for row in all_refs])

    valid = False

    if mandate:
        if uid and sig:
            valid = (
                uid == mandate["mandate_uid"]
                and verify_signature(reference, uid, sig)
            )
        else:
            # Si ref seule est fournie, on montre le mandat trouvé sans valider la signature.
            valid = True

    return render_template("verify.html", mandate=mandate, valid=valid, ref=reference)


@app.route("/documents/<kind>/<path:reference>")
@login_required
def open_document(kind, reference):
    mandate = fetch_one("SELECT * FROM mandates WHERE reference = ?", (reference,))
    if not mandate:
        abort(404)

    if kind == "pdf":
        file_path = mandate["pdf_path"]
    elif kind == "docx":
        file_path = mandate["docx_path"]
    elif kind == "qr":
        file_path = mandate["qr_path"]
    else:
        abort(404)

    if not file_path or not Path(file_path).exists():
        abort(404)

    return send_file(file_path, as_attachment=False)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)
