from pathlib import Path
import os

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
DOCX_DIR = OUTPUT_DIR / "docx"
PDF_DIR = OUTPUT_DIR / "pdf"
QR_DIR = OUTPUT_DIR / "qrcodes"
STATIC_DIR = BASE_DIR / "static"

DB_PATH = DATA_DIR / "registry.db"

SECRET_KEY = os.getenv("MANDATE_SECRET_KEY", "CHANGE_ME_TO_A_LONG_RANDOM_SECRET")
FLASK_SECRET = "RN_SUPER_SECRET_2026_ULTRA_STABLE"

ORG_NAME = "ONG Renaître de Nouveau"
ORG_SHORT = "ONG-RN"
ORG_WEBSITE = os.getenv("ORG_WEBSITE", "https://www.renaitredenouveau.org")
ORG_EMAIL = os.getenv("ORG_EMAIL", "contact@renaitredenouveau.org")
ORG_PHONE = os.getenv("ORG_PHONE", "+33 7 51 40 71 88")
ORG_CITY = os.getenv("ORG_CITY", "Strasbourg")
ORG_COUNTRY = os.getenv("ORG_COUNTRY", "France")
REGISTRE_ASSOCIATION = os.getenv("REGISTRE_ASSOCIATION", "À compléter avec les références officielles de l'association")

PRESIDENT_NAME = os.getenv("PRESIDENT_NAME", "Romuald HOUNYEME")
PRESIDENT_TITLE = "Président du Conseil d’Administration (PCA)"
COORDINATOR_TITLE = "Superviseur / Coordinateur national au Bénin"

BASE_VERIFY_URL = os.getenv("BASE_VERIFY_URL", "https://verify.renaitredenouveau.org/verify")

ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "admin123"

DEFAULT_STATUSES = ["actif", "suspendu", "revoque", "expire"]

LOGO_PATH = STATIC_DIR / "logo.png"
PRESIDENT_SIGNATURE_PATH = STATIC_DIR / "signature_president.png"
STAMP_PATH = STATIC_DIR / "cachet.png"
