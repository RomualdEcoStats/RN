import csv
import argparse
from pathlib import Path

from database import init_db, insert_mandate, fetch_one
from security import generate_uid, generate_signature
from generator import generate_all

def auto_reference(prefix: str, year: str, idx: int) -> str:
    return f"{prefix}/{year}/ONG-RN/PCA/CSA/{idx:04d}"

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv", required=True, help="Chemin du CSV")
    args = parser.parse_args()

    init_db()
    csv_path = Path(args.csv)
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        created = 0
        for idx, row in enumerate(reader, start=1):
            reference = row.get("reference") or auto_reference(
                row.get("reference_prefix", "RN"),
                (row.get("date_emission") or "2026")[:4],
                idx
            )
            if fetch_one("SELECT * FROM mandates WHERE reference=?", (reference,)):
                continue
            uid = generate_uid()
            sig = generate_signature(reference, uid)
            payload = {
                "reference": reference,
                "mandate_uid": uid,
                "signature_token": sig,
                "civilite": row.get("civilite",""),
                "nom": row.get("nom",""),
                "prenom": row.get("prenom",""),
                "adresse": row.get("adresse",""),
                "telephone": row.get("telephone",""),
                "email": row.get("email",""),
                "fonction": row.get("fonction","Délégué"),
                "zone_intervention": row.get("zone_intervention",""),
                "date_emission": row.get("date_emission",""),
                "date_expiration": row.get("date_expiration",""),
                "statut": row.get("statut","actif"),
                "ville_signature": row.get("ville_signature","Strasbourg"),
                "pays": row.get("pays","Bénin"),
                "preferences_affichage": row.get("preferences_affichage",""),
                "notes": row.get("notes",""),
            }
            qr, docx, pdf = generate_all(payload)
            payload["qr_path"] = str(qr)
            payload["docx_path"] = str(docx)
            payload["pdf_path"] = str(pdf)
            insert_mandate(payload)
            created += 1
    print(f"{created} mandat(s) importé(s).")

if __name__ == "__main__":
    main()
