from pathlib import Path
from urllib.parse import urlencode

import qrcode
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from config import (
    DOCX_DIR,
    PDF_DIR,
    QR_DIR,
    BASE_VERIFY_URL,
    ORG_NAME,
    ORG_WEBSITE,
    ORG_EMAIL,
    ORG_PHONE,
    ORG_CITY,
    PRESIDENT_NAME,
    PRESIDENT_TITLE,
    REGISTRE_ASSOCIATION,
    COORDINATOR_TITLE,
    LOGO_PATH,
    PRESIDENT_SIGNATURE_PATH,
    STAMP_PATH,
)

BODY_INTRO = (
    "Nous soussigné {org_name}, inscrite au Registre des Associations {registre}, "
    "par {president}, agissant en qualité de {president_title}, attestons par la présente que "
    "{civilite_nom}, demeurant à {adresse}, contact {telephone}, membre de {org_name}, est autorisé(e) "
    "à représenter cette ONG en qualité de {fonction} dans la zone d’intervention {zone} qui lui est fixée au {pays}."
)

BODY_MIDDLE = (
    "A ce titre, sur la base des objectifs définis dans les statuts de l’ONG et conformément au plan "
    "de travail qui lui sera soumis par l’intermédiaire du {coord_title}, il/elle devra :"
)

MISSIONS = [
    "Agir dans sa zone d’intervention au nom de l’ONG dans le cadre de sa mission ;",
    "Assurer la mise en œuvre des activités sociales, humanitaires et communautaires dans sa zone d’intervention ;",
    "Initier, coordonner et accompagner des actions et projets de l’ONG en liaison avec le PCA et le Superviseur / Coordinateur national ;",
    "Veiller rigoureusement à ce que l’assistance atteigne les vraies cibles selon les prévisions des projets mis en œuvre ;",
    "Collaborer avec les autorités locales, les institutions et les partenaires au développement pour faciliter l’accomplissement de la mission de l’ONG sur le terrain ;",
    "Rendre compte périodiquement au PCA et au Coordinateur national des tâches accomplies et formuler des propositions face aux défis rencontrés.",
]

BODY_END = [
    "Le Délégué de l’ONG reconnait exécuter sa mission en qualité de bénévole et s’engage à ne pas réclamer de rémunération pour ses prestations sur le terrain. Toutefois, l’ONG reconnait devoir prendre en charge toutes ses dépenses liées à l’exercice de sa mission, telles que déplacement, communications téléphoniques, reportage d’évènement, etc.",
    "Le Délégué de l’ONG accepte par ailleurs de ne prendre, au nom de l’ONG, aucune initiative en dehors des objectifs et du plan de travail de l’ONG.",
    "Le présent mandat peut être, en cas de besoin, modifié ou résilié sur demande de l’une des deux parties.",
    "En foi de quoi, il est délivré pour servir et valoir ce que de droit.",
]


def sanitize_reference(ref: str) -> str:
    return ref.replace("/", "_")


def verify_url(reference: str, uid: str, sig: str) -> str:
    """
    Construit une URL de vérification robuste en encodant correctement
    tous les paramètres de requête.
    """
    query = urlencode(
        {
            "ref": reference,
            "uid": uid,
            "sig": sig,
        }
    )
    return f"{BASE_VERIFY_URL}?{query}"


def generate_qr(reference: str, uid: str, sig: str) -> Path:
    QR_DIR.mkdir(parents=True, exist_ok=True)
    path = QR_DIR / f"{sanitize_reference(reference)}_qr.png"
    img = qrcode.make(verify_url(reference, uid, sig))
    img.save(path)
    return path


def build_text(payload: dict) -> dict:
    civilite_nom = f"{payload.get('civilite', '').strip()} {payload['prenom']} {payload['nom']}".strip()
    fmt = {
        "org_name": ORG_NAME,
        "registre": REGISTRE_ASSOCIATION,
        "president": PRESIDENT_NAME,
        "president_title": PRESIDENT_TITLE,
        "civilite_nom": civilite_nom,
        "adresse": payload.get("adresse") or "Non renseignée",
        "telephone": payload.get("telephone") or "Non renseigné",
        "fonction": payload["fonction"],
        "zone": payload["zone_intervention"],
        "pays": payload.get("pays") or "Bénin",
        "coord_title": COORDINATOR_TITLE,
    }
    return fmt


def generate_docx(payload: dict, qr_path: Path) -> Path:
    DOCX_DIR.mkdir(parents=True, exist_ok=True)
    ref = payload["reference"]
    out = DOCX_DIR / f"{sanitize_reference(ref)}.docx"

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Mm(14)
    section.bottom_margin = Mm(14)
    section.left_margin = Mm(18)
    section.right_margin = Mm(18)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if LOGO_PATH.exists():
        p.add_run().add_picture(str(LOGO_PATH), width=Mm(25))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(ORG_NAME.upper())
    r.bold = True
    r.font.size = Pt(16)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ATTESTATION DE MANDAT")
    r.bold = True
    r.font.size = Pt(18)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"N° {ref}").bold = True

    fmt = build_text(payload)

    doc.add_paragraph(BODY_INTRO.format(**fmt))
    doc.add_paragraph(BODY_MIDDLE.format(**fmt))

    for item in MISSIONS:
        doc.add_paragraph(item, style=None).style = doc.styles["List Bullet"]

    for item in BODY_END:
        doc.add_paragraph(item)

    doc.add_paragraph(f"Date d’émission : {payload.get('date_emission', '')}")
    doc.add_paragraph(f"Date d’expiration : {payload.get('date_expiration', '')}")
    doc.add_paragraph(f"Statut : {payload.get('statut', 'actif')}")

    doc.add_paragraph(
        f"Vérification : {verify_url(payload['reference'], payload['mandate_uid'], payload['signature_token'])}"
    )

    p = doc.add_paragraph()
    if qr_path.exists():
        p.add_run().add_picture(str(qr_path), width=Mm(35))
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    table = doc.add_table(rows=2, cols=3)

    hdr = table.rows[0].cells
    hdr[0].text = f"{payload.get('ville_signature', 'Strasbourg')}, le {payload.get('date_emission', '')}"
    hdr[1].text = "Signature du Président"
    hdr[2].text = "Signature du délégué"

    row = table.rows[1].cells
    row[0].text = "Cachet ONG"

    if STAMP_PATH.exists():
        row[0].paragraphs[0].add_run().add_picture(str(STAMP_PATH), width=Mm(25))

    if PRESIDENT_SIGNATURE_PATH.exists():
        row[1].paragraphs[0].add_run().add_picture(str(PRESIDENT_SIGNATURE_PATH), width=Mm(35))
        row[1].add_paragraph(PRESIDENT_NAME)
        row[1].add_paragraph(PRESIDENT_TITLE)

    row[2].text = f"{payload['prenom']} {payload['nom']}\nSignature précédée de la mention lu et approuvé"

    doc.save(out)
    return out


def generate_pdf(payload: dict, qr_path: Path) -> Path:
    PDF_DIR.mkdir(parents=True, exist_ok=True)
    ref = payload["reference"]
    out = PDF_DIR / f"{sanitize_reference(ref)}.pdf"

    c = canvas.Canvas(str(out), pagesize=A4)
    w, h = A4

    if LOGO_PATH.exists():
        try:
            c.drawImage(
                ImageReader(str(LOGO_PATH)),
                25 * mm,
                h - 40 * mm,
                width=22 * mm,
                height=22 * mm,
                preserveAspectRatio=True,
                mask="auto",
            )
            c.saveState()
            c.setFillAlpha(0.08)
            c.drawImage(
                ImageReader(str(LOGO_PATH)),
                60 * mm,
                90 * mm,
                width=90 * mm,
                height=90 * mm,
                preserveAspectRatio=True,
                mask="auto",
            )
            c.restoreState()
        except Exception:
            pass

    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(w / 2, h - 25 * mm, "ATTESTATION DE MANDAT")

    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(w / 2, h - 32 * mm, f"N° {ref}")

    y = h - 45 * mm
    c.setFont("Helvetica", 10)

    fmt = build_text(payload)
    paragraphs = [
        BODY_INTRO.format(**fmt),
        BODY_MIDDLE.format(**fmt),
        *[f"• {m}" for m in MISSIONS],
        *BODY_END,
        f"Date d’émission : {payload.get('date_emission', '')}",
        f"Date d’expiration : {payload.get('date_expiration', '')}",
        f"Statut au moment de l’édition : {payload.get('statut', 'actif')}",
        f"Site officiel : {ORG_WEBSITE}",
        f"Contact : {ORG_EMAIL} | {ORG_PHONE}",
    ]

    for para in paragraphs:
        for line in _wrap_text(para, 105):
            c.drawString(20 * mm, y, line)
            y -= 5.2 * mm
            if y < 55 * mm:
                c.showPage()
                y = h - 20 * mm
                c.setFont("Helvetica", 10)

    if qr_path.exists():
        try:
            c.drawImage(
                ImageReader(str(qr_path)),
                20 * mm,
                18 * mm,
                width=28 * mm,
                height=28 * mm,
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass

    c.setFont("Helvetica", 8)
    verif = verify_url(payload["reference"], payload["mandate_uid"], payload["signature_token"])
    for i, line in enumerate(_wrap_text("Vérification numérique : " + verif, 90)):
        c.drawString(55 * mm, 35 * mm - i * 4 * mm, line)

    if STAMP_PATH.exists():
        try:
            c.drawImage(
                ImageReader(str(STAMP_PATH)),
                115 * mm,
                18 * mm,
                width=25 * mm,
                height=25 * mm,
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass

    if PRESIDENT_SIGNATURE_PATH.exists():
        try:
            c.drawImage(
                ImageReader(str(PRESIDENT_SIGNATURE_PATH)),
                145 * mm,
                18 * mm,
                width=35 * mm,
                height=15 * mm,
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass

    c.setFont("Helvetica-Bold", 9)
    c.drawString(115 * mm, 15 * mm, "Cachet ONG")
    c.drawString(145 * mm, 15 * mm, PRESIDENT_NAME)
    c.drawString(145 * mm, 11 * mm, PRESIDENT_TITLE)

    c.save()
    return out


def _wrap_text(text: str, width: int) -> list[str]:
    words = text.split()
    lines = []
    current = []
    count = 0

    for word in words:
        extra = 1 if current else 0
        if count + len(word) + extra <= width:
            current.append(word)
            count += len(word) + extra
        else:
            lines.append(" ".join(current))
            current = [word]
            count = len(word)

    if current:
        lines.append(" ".join(current))

    return lines


def generate_all(payload: dict):
    qr = generate_qr(payload["reference"], payload["mandate_uid"], payload["signature_token"])
    docx = generate_docx(payload, qr)
    pdf = generate_pdf(payload, qr)
    return qr, docx, pdf