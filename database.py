import sqlite3
from config import DB_PATH, DATA_DIR

SCHEMA = '''
CREATE TABLE IF NOT EXISTS mandates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    reference TEXT UNIQUE NOT NULL,
    mandate_uid TEXT UNIQUE NOT NULL,
    signature_token TEXT NOT NULL,
    civilite TEXT,
    nom TEXT NOT NULL,
    prenom TEXT NOT NULL,
    adresse TEXT,
    telephone TEXT,
    email TEXT,
    fonction TEXT NOT NULL,
    zone_intervention TEXT NOT NULL,
    date_emission TEXT NOT NULL,
    date_expiration TEXT,
    statut TEXT NOT NULL DEFAULT 'actif',
    ville_signature TEXT DEFAULT 'Strasbourg',
    pays TEXT DEFAULT 'Bénin',
    preferences_affichage TEXT,
    notes TEXT,
    photo_path TEXT,
    logo_path TEXT,
    qr_path TEXT,
    docx_path TEXT,
    pdf_path TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);
'''

def init_db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_PATH) as conn:
        conn.executescript(SCHEMA)
        conn.commit()

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def fetch_all(query, params=()):
    with get_conn() as conn:
        return conn.execute(query, params).fetchall()

def fetch_one(query, params=()):
    with get_conn() as conn:
        return conn.execute(query, params).fetchone()

def execute(query, params=()):
    with get_conn() as conn:
        cur = conn.execute(query, params)
        conn.commit()
        return cur.rowcount

def insert_mandate(payload: dict):
    keys = sorted(payload.keys())
    cols = ", ".join(keys)
    placeholders = ", ".join(["?" for _ in keys])
    values = [payload[k] for k in keys]
    with get_conn() as conn:
        conn.execute(f"INSERT INTO mandates ({cols}) VALUES ({placeholders})", values)
        conn.commit()

def update_mandate(reference: str, payload: dict):
    keys = sorted(payload.keys())
    setters = ", ".join([f"{k}=?" for k in keys])
    values = [payload[k] for k in keys] + [reference]
    with get_conn() as conn:
        cur = conn.execute(
            f"UPDATE mandates SET {setters}, updated_at=CURRENT_TIMESTAMP WHERE reference=?",
            values,
        )
        conn.commit()
        return cur.rowcount
