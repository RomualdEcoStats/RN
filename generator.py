from pathlib import Path
from urllib.parse import urlencode

import qrcode
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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
from security import generate_signature


# -------------------------------------------------------------------
# Références institutionnelles fixes
# -------------------------------------------------------------------

AMALIA_REFERENCE = "A2019STR000236 – Volume 97 – Folio 229"
ORG_HEAD_OFFICE = "24 Rue de la Niederbourg, 67400 Illkirch-Graffenstaden"

TITLE_TEXT = "ATTESTATION DE MANDAT OFFICIEL"

DECLARATION_TEXT = (
    "La présente attestation est délivrée pour servir et valoir ce que de droit "
    "auprès de toute autorité administrative, institutionnelle ou partenaire."
)

# -------------------------------------------------------------------
# Texte ministériel / diplomatique
# -------------------------------------------------------------------

INTRO_BLOCK = (
    "L’ONG Renaître de Nouveau, organisation régulièrement constituée et inscrite "
    f"au Registre des Associations sous la référence AMALIA {AMALIA_REFERENCE}, "
    f"ayant son siège au {ORG_HEAD_OFFICE}, représentée par Monsieur {PRESIDENT_NAME}, "
    "agissant en qualité de Président du Conseil d’Administration (PCA), "
    "certifie et atteste que :"
)

MANDATE_IDENTITY_BLOCK = (
    "{civilite_nom}, demeurant à {adresse}, contact : {telephone}, membre de ladite organisation, "
    "est dûment habilité à représenter l’ONG Renaître de Nouveau en qualité de :"
)

MANDATE_TITLE_BLOCK = "{fonction}"

MANDATE_ZONE_BLOCK = (
    "dans la zone d’intervention suivante : {zone_label}."
)

ARTICLE_1_TITLE = "1. Portée du mandat"
ARTICLE_1_TEXT = (
    "Le présent mandat constitue une habilitation officielle conférée par l’ONG, "
    "autorisant son titulaire à intervenir en son nom dans le cadre strict de ses missions, "
    "conformément :"
)
ARTICLE_1_BULLETS = [
    "aux statuts de l’organisation ;",
    "aux orientations stratégiques validées ;",
    "et aux instructions des instances dirigeantes.",
]

ARTICLE_2_TITLE = "2. Attributions"
ARTICLE_2_TEXT = "Dans ce cadre, le titulaire est autorisé à :"
ARTICLE_2_BULLETS = [
    "représenter l’ONG auprès des autorités administratives, institutions et partenaires ;",
    "contribuer à la mise en œuvre des activités sociales, humanitaires et communautaires ;",
    "assurer la coordination locale des actions validées par l’organisation ;",
    "rendre compte des activités menées conformément aux procédures internes.",
]

ARTICLE_3_TITLE = "3. Limites de l’habilitation"
ARTICLE_3_TEXT_1 = (
    "Le présent mandat n’emporte aucun pouvoir d’engagement juridique, financier ou contractuel "
    "autonome au nom de l’ONG, sauf autorisation expresse et préalable."
)
ARTICLE_3_TEXT_2 = "Le titulaire est tenu d’agir dans le strict respect :"
ARTICLE_3_BULLETS = [
    "des lois et règlements en vigueur dans le pays d’intervention ;",
    "des statuts et directives de l’ONG.",
]

ARTICLE_4_TITLE = "4. Nature de la mission"
ARTICLE_4_TEXT = (
    "La mission est exercée à titre bénévole. Toute prise en charge de frais éventuels "
    "est soumise aux procédures internes de l’ONG."
)

ARTICLE_5_TITLE = "5. Durée et révocation"
ARTICLE_5_TEXT = (
    "Le présent mandat est valable du {date_emission} au {date_expiration}, sauf modification, "
    "suspension ou retrait à tout moment par l’ONG, notamment en cas de manquement."
)

ARTICLE_6_TITLE = "6. Authenticité"
ARTICLE_6_TEXT = (
    "Le présent document est sécurisé par un dispositif de vérification numérique "
    "(QR code / lien officiel) permettant d’en confirmer l’authenticité et la validité."
)


# -------------------------------------------------------------------
# Helpers de style DOCX
# -------------------------------------------------------------------

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


def _set_cell_shading(cell, fill: str = "D9EAF7"):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def _set_table_borders(table, color: str = "1F4E79", size: str = "8"):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        elem = OxmlElement(f"w:{edge}")
        elem.set(qn("w:val"), "single")
        elem.set(qn("w:sz"), size)
        elem.set(qn("w:space"), "0")
        elem.set(qn("w:color"), color)
        borders.append(elem)
    tbl_pr.append(borders)


def _add_bullet(doc: Document, text: str):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    run = p.add_run(text)
    run.font.size = Pt(10.2)


def _add_docx_heading(doc: Document, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(10.8)
    run.font.color.rgb = RGBColor(31, 78, 121)


def _add_docx_paragraph(doc: Document, text: str, size: float = 10.2, justify: bool = True):
    p = doc.add_paragraph()
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(text)
    run.font.size = Pt(size)
    return p


# -------------------------------------------------------------------
# Génération QR
# -------------------------------------------------------------------

def generate_qr(reference: str, uid: str, sig: str) -> Path:
    QR_DIR.mkdir(parents=True, exist_ok=True)
    path = QR_DIR / f"{sanitize_reference(reference)}_qr.png"
    qr_url = verify_url(reference, uid, sig)
    img = qrcode.make(qr_url)
    img.save(path)
    return path


# -------------------------------------------------------------------
# Préparation du texte dynamique
# -------------------------------------------------------------------

def build_text(payload: dict) -> dict:
    civilite_nom = f"{payload.get('civilite', '').strip()} {payload['prenom']} {payload['nom']}".strip()

    zone = payload.get("zone_intervention", "").strip()
    pays = payload.get("pays") or "Bénin"

    # Construction plus institutionnelle de la zone
    if zone and pays:
        zone_label = f"{zone}, République du {pays}" if pays.lower() != "bénin" else f"Département de l’{zone}, République du Bénin"
        if zone.lower().startswith(("atacora", "ouémé", "alibori", "borgou", "mono", "zou", "couffo", "plateau", "collines", "donga", "littoral", "atakora")):
            zone_label = f"Département de l’{zone}, République du Bénin" if zone[0].lower() in "aeiouyh" else f"Département du {zone}, République du Bénin"
    else:
        zone_label = f"{zone} ({pays})" if zone else pays

    fmt = {
        "civilite_nom": civilite_nom,
        "adresse": payload.get("adresse") or "Non renseignée",
        "telephone": payload.get("telephone") or "Non renseigné",
        "fonction": payload["fonction"],
        "zone_label": zone_label,
        "date_emission": payload.get("date_emission") or "[date]",
        "date_expiration": payload.get("date_expiration") or "[date]",
    }
    return fmt


# -------------------------------------------------------------------
# DOCX
# -------------------------------------------------------------------

def generate_docx(payload: dict, qr_path: Path) -> Path:
    DOCX_DIR.mkdir(parents=True, exist_ok=True)
    ref = payload["reference"]
    out = DOCX_DIR / f"{sanitize_reference(ref)}.docx"

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Mm(10)
    section.bottom_margin = Mm(10)
    section.left_margin = Mm(16)
    section.right_margin = Mm(16)

    # En-tête logo + organisation
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if LOGO_PATH.exists():
        p.add_run().add_picture(str(LOGO_PATH), width=Mm(22))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(ORG_NAME.upper())
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = RGBColor(31, 78, 121)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(ORG_HEAD_OFFICE)
    run.font.size = Pt(9)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"Référence AMALIA : {AMALIA_REFERENCE}")
    run.font.size = Pt(9)
    run.bold = True

    # Titre encadré en couleur
    title_table = doc.add_table(rows=1, cols=1)
    title_table.autofit = True
    _set_table_borders(title_table, color="1F4E79", size="10")
    title_cell = title_table.cell(0, 0)
    _set_cell_shading(title_cell, fill="D9EAF7")
    p = title_cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(TITLE_TEXT)
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(31, 78, 121)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"N° {ref}")
    run.bold = True
    run.font.size = Pt(10.5)

    fmt = build_text(payload)

    _add_docx_paragraph(doc, INTRO_BLOCK, size=10.2, justify=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("Certifie et atteste que :")
    run.bold = True
    run.font.size = Pt(10.4)

    _add_docx_paragraph(doc, MANDATE_IDENTITY_BLOCK.format(**fmt), size=10.2, justify=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(MANDATE_TITLE_BLOCK.format(**fmt))
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 78, 121)

    _add_docx_paragraph(doc, MANDATE_ZONE_BLOCK.format(**fmt), size=10.2, justify=True)

    _add_docx_heading(doc, ARTICLE_1_TITLE)
    _add_docx_paragraph(doc, ARTICLE_1_TEXT, size=10.0, justify=True)
    for item in ARTICLE_1_BULLETS:
        _add_bullet(doc, item)

    _add_docx_heading(doc, ARTICLE_2_TITLE)
    _add_docx_paragraph(doc, ARTICLE_2_TEXT, size=10.0, justify=True)
    for item in ARTICLE_2_BULLETS:
        _add_bullet(doc, item)

    _add_docx_heading(doc, ARTICLE_3_TITLE)
    _add_docx_paragraph(doc, ARTICLE_3_TEXT_1, size=10.0, justify=True)
    _add_docx_paragraph(doc, ARTICLE_3_TEXT_2, size=10.0, justify=True)
    for item in ARTICLE_3_BULLETS:
        _add_bullet(doc, item)

    _add_docx_heading(doc, ARTICLE_4_TITLE)
    _add_docx_paragraph(doc, ARTICLE_4_TEXT, size=10.0, justify=True)

    _add_docx_heading(doc, ARTICLE_5_TITLE)
    _add_docx_paragraph(doc, ARTICLE_5_TEXT.format(**fmt), size=10.0, justify=True)

    _add_docx_heading(doc, ARTICLE_6_TITLE)
    _add_docx_paragraph(doc, ARTICLE_6_TEXT, size=10.0, justify=True)

    _add_docx_heading(doc, "Déclaration")
    _add_docx_paragraph(doc, DECLARATION_TEXT, size=10.0, justify=True)

    # Bloc vérification
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(
        f"Vérification : {verify_url(payload['reference'], payload['mandate_uid'], payload['signature_token'])}"
    )
    run.font.size = Pt(8.2)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if qr_path.exists():
        p.add_run().add_picture(str(qr_path), width=Mm(26))

    # Bloc signature / cachet - un seul signataire
    sign_table = doc.add_table(rows=2, cols=2)
    _set_table_borders(sign_table, color="7F9DB9", size="6")

    hdr = sign_table.rows[0].cells
    hdr[0].text = f"Fait à {payload.get('ville_signature', ORG_CITY)}, le {payload.get('date_emission', '[date]')}"
    hdr[1].text = "Pour l’ONG Renaître de Nouveau\nLe Président du Conseil d’Administration"

    row = sign_table.rows[1].cells
    row[0].text = "Cachet officiel"
    if STAMP_PATH.exists():
        row[0].paragraphs[0].add_run().add_picture(str(STAMP_PATH), width=Mm(24))

    row[1].text = ""
    if PRESIDENT_SIGNATURE_PATH.exists():
        row[1].paragraphs[0].add_run().add_picture(str(PRESIDENT_SIGNATURE_PATH), width=Mm(34))
    row[1].add_paragraph(PRESIDENT_NAME)
    row[1].add_paragraph(PRESIDENT_TITLE)

    doc.save(out)
    return out


# -------------------------------------------------------------------
# PDF institutionnel
# -------------------------------------------------------------------

def generate_pdf(payload: dict, qr_path: Path) -> Path:
    PDF_DIR.mkdir(parents=True, exist_ok=True)
    ref = payload["reference"]
    out = PDF_DIR / f"{sanitize_reference(ref)}.pdf"

    c = canvas.Canvas(str(out), pagesize=A4)
    w, h = A4

    # Fond / filigrane logo
    if LOGO_PATH.exists():
        try:
            c.saveState()
            c.setFillAlpha(0.08)
            c.drawImage(
                ImageReader(str(LOGO_PATH)),
                45 * mm,
                70 * mm,
                width=120 * mm,
                height=120 * mm,
                preserveAspectRatio=True,
                mask="auto",
            )
            c.restoreState()
        except Exception:
            pass

    # Logo haut de page
    if LOGO_PATH.exists():
        try:
            c.drawImage(
                ImageReader(str(LOGO_PATH)),
                18 * mm,
                h - 30 * mm,
                width=18 * mm,
                height=18 * mm,
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass

    # En-tête institutionnel
    c.setFont("Helvetica-Bold", 13)
    c.setFillColor(colors.HexColor("#1F4E79"))
    c.drawCentredString(w / 2, h - 18 * mm, ORG_NAME.upper())

    c.setFillColor(colors.black)
    c.setFont("Helvetica", 9)
    c.drawCentredString(w / 2, h - 23 * mm, ORG_HEAD_OFFICE)
    c.drawCentredString(w / 2, h - 27 * mm, f"Référence AMALIA : {AMALIA_REFERENCE}")

    # Titre encadré en couleur
    box_x = 20 * mm
    box_y = h - 40 * mm
    box_w = w - 40 * mm
    box_h = 10 * mm
    c.setFillColor(colors.HexColor("#D9EAF7"))
    c.setStrokeColor(colors.HexColor("#1F4E79"))
    c.setLineWidth(1.2)
    c.roundRect(box_x, box_y, box_w, box_h, 2 * mm, stroke=1, fill=1)

    c.setFillColor(colors.HexColor("#1F4E79"))
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(w / 2, box_y + 3.2 * mm, TITLE_TEXT)

    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(w / 2, h - 46 * mm, f"N° {ref}")

    fmt = build_text(payload)

    y = h - 55 * mm
    c.setFont("Helvetica", 9.1)

    text_blocks = [
        INTRO_BLOCK,
        "Certifie et atteste que :",
        MANDATE_IDENTITY_BLOCK.format(**fmt),
        MANDATE_TITLE_BLOCK.format(**fmt),
        MANDATE_ZONE_BLOCK.format(**fmt),
        ARTICLE_1_TITLE,
        ARTICLE_1_TEXT,
        *[f"• {x}" for x in ARTICLE_1_BULLETS],
        ARTICLE_2_TITLE,
        ARTICLE_2_TEXT,
        *[f"• {x}" for x in ARTICLE_2_BULLETS],
        ARTICLE_3_TITLE,
        ARTICLE_3_TEXT_1,
        ARTICLE_3_TEXT_2,
        *[f"• {x}" for x in ARTICLE_3_BULLETS],
        ARTICLE_4_TITLE,
        ARTICLE_4_TEXT,
        ARTICLE_5_TITLE,
        ARTICLE_5_TEXT.format(**fmt),
        ARTICLE_6_TITLE,
        ARTICLE_6_TEXT,
        "Déclaration",
        DECLARATION_TEXT,
    ]

    for block in text_blocks:
        is_heading = block in {
            ARTICLE_1_TITLE,
            ARTICLE_2_TITLE,
            ARTICLE_3_TITLE,
            ARTICLE_4_TITLE,
            ARTICLE_5_TITLE,
            ARTICLE_6_TITLE,
            "Déclaration",
            "Certifie et atteste que :",
        }

        if is_heading:
            c.setFont("Helvetica-Bold", 9.5)
            c.setFillColor(colors.HexColor("#1F4E79"))
        elif block == MANDATE_TITLE_BLOCK.format(**fmt):
            c.setFont("Helvetica-Bold", 10.5)
            c.setFillColor(colors.HexColor("#1F4E79"))
        else:
            c.setFont("Helvetica", 9.1)
            c.setFillColor(colors.black)

        lines = _wrap_text(block, 115)
        for line in lines:
            if block == MANDATE_TITLE_BLOCK.format(**fmt):
                c.drawCentredString(w / 2, y, line)
            else:
                c.drawString(18 * mm, y, line)
            y -= 4.2 * mm

            if y < 42 * mm:
                c.showPage()
                y = h - 20 * mm
                c.setFont("Helvetica", 9.1)
                c.setFillColor(colors.black)

    # QR + lien
    if qr_path.exists():
        try:
            c.drawImage(
                ImageReader(str(qr_path)),
                18 * mm,
                14 * mm,
                width=24 * mm,
                height=24 * mm,
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass

    verif = verify_url(payload["reference"], payload["mandate_uid"], payload["signature_token"])
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 7.5)
    for i, line in enumerate(_wrap_text("Vérification numérique : " + verif, 92)):
        c.drawString(46 * mm, 30 * mm - i * 3.5 * mm, line)

    # Bloc signature / cachet
    if STAMP_PATH.exists():
        try:
            c.drawImage(
                ImageReader(str(STAMP_PATH)),
                118 * mm,
                14 * mm,
                width=24 * mm,
                height=24 * mm,
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
                17 * mm,
                width=38 * mm,
                height=14 * mm,
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass

    c.setFont("Helvetica", 8.2)
    c.drawString(118 * mm, 11 * mm, "Cachet officiel")
    c.drawString(145 * mm, 11 * mm, PRESIDENT_NAME)
    c.drawString(145 * mm, 7 * mm, PRESIDENT_TITLE)
    c.drawString(18 * mm, 11 * mm, f"Fait à {payload.get('ville_signature', ORG_CITY)}, le {payload.get('date_emission', '[date]')}")

    c.save()
    return out


# -------------------------------------------------------------------
# Wrap helper
# -------------------------------------------------------------------

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


# -------------------------------------------------------------------
# Génération globale
# -------------------------------------------------------------------

def generate_all(payload: dict):
    # Recalcul systématique de la signature pour garantir
    # une parfaite cohérence entre QR, PDF et DOCX.
    payload["signature_token"] = generate_signature(
        payload["reference"],
        payload["mandate_uid"],
    )

    qr = generate_qr(payload["reference"], payload["mandate_uid"], payload["signature_token"])
    docx = generate_docx(payload, qr)
    pdf = generate_pdf(payload, qr)
    return qr, docx, pdf