#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VDM Prijslijst Studio — Lokale Engine, Markdown Mailer & Sanitizer (v1.2.0)
100% Offline, Pure Python Stdlib (Geen externe libraries nodig).
Ontwikkeld voor Vandermaesen NV (Nick Cuppers) door Webbit / Robin Stroobants.
"""

import os
import re
import io
import csv
import json
import email
import ssl
import urllib.request
from email import policy
from email.message import EmailMessage
from email.utils import formatdate
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta

APP_VERSION = "1.2.0"
# Zodra je de GitHub repo aanmaakt, vul je hier de raw base URL in voor 1-klik updates bij Nick:
GITHUB_RAW_BASE_URL = os.environ.get(
    "VDM_UPDATE_URL",
    "https://raw.githubusercontent.com/robinstroobants/vandermaesen-studio/main"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIDENTIAL_DIR = os.path.join(BASE_DIR, "confidential")
RAW_DIR = CONFIDENTIAL_DIR
DATA_DIR = os.path.join(CONFIDENTIAL_DIR, "data")
OUT_DIR = os.path.join(CONFIDENTIAL_DIR, "output")
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates", "mails")

LOGO_PATH = os.path.join(ASSETS_DIR, "vdm_logo.png")
BRON_NL_PATH = os.path.join(CONFIDENTIAL_DIR, "Art. Prijzen_BRON_20260601.xlsx")
VOORWAARDEN_PATH = os.path.join(CONFIDENTIAL_DIR, "Overzicht speciale voorwaarden klanten 121125.xlsx")
CONFIG_JSON_PATH = os.path.join(DATA_DIR, "klanten_config.json")
SETTINGS_JSON_PATH = os.path.join(DATA_DIR, "campagne_settings.json")
MASTER_EXCEL_PATH = os.path.join(CONFIDENTIAL_DIR, "00_MASTER_KLANTENCONFIG.xlsx")

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
ET.register_namespace("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")

NL_MAANDEN = [
    "", "januari", "februari", "maart", "april", "mei", "juni",
    "juli", "augustus", "september", "oktober", "november", "december"
]
FR_MAANDEN = [
    "", "janvier", "février", "mars", "avril", "mai", "juin",
    "juillet", "août", "septembre", "octobre", "novembre", "décembre"
]

STANDARD_TABS_META = [
    ("RP", "Ronde palen (RP)"),
    ("SP", "Steunpalen (SP)"),
    ("TOEBEH", "Toebehoren (TOEBEH)"),
    ("DS 2D", "Dubbel staalmatten 2D (DS 2D)"),
    ("PREST 3D", "Prestipli staalmatten 3D (PREST 3D)"),
    ("VPI", "Vierkante palen met inserts (VPI)"),
    ("VPZ", "Vierkante palen (VPZ)"),
    ("VPUS", "Vierkante palen met smalle U (VPUS)"),
    ("VPUB", "Vierkante palen met brede U (VPUB)"),
    ("VP + VOET", "Palen op voetplaat + losse voeten (VP+VOET)"),
    ("SKP", "Steenkorfpalen (SKP)"),
    ("VDM GP", "VDM gleufpalen (VDM GP)"),
    ("VLP GP", "VLP gleufpalen (VLP GP)"),
    ("MEHA GP", "MEHA gleufpalen (MEHA GP)"),
    ("VDM BAND", "VDM band (VDM BAND)"),
    ("BPH", "Betonplaathouders (BPH)"),
    ("EP", "Enkele tuinpoorten (EP)"),
    ("DP", "Dubbele tuinpoorten (DP)"),
    ("EPDS", "Enkele poorten met dubbel staalmatvulling (EPDS)"),
    ("DPDS", "Dubbele poorten met dubbel staalmatvulling (DPDS)"),
    ("MAATP", "Toeslag maatpoorten (MAATP)"),
]

OPTIONAL_TABS_META = [
    ("PK", "Poortkaders (PK)"),
    ("PK DS", "Poortkaders dubbel staalmat (PK DS)"),
    ("PP SETS", "Sets poortpalen (PP SETS)"),
    ("SLOTEN", "Sloten (SLOTEN)"),
    ("RP SCHETS", "Ronde palen met schetsplaat (RP SCHETS)"),
    ("RP PRES", "Ronde palen met geperst uiteinde (RP PRES)"),
    ("VPB", "Vierkante palen met boringen (VPB)"),
    ("RP VERPAKT", "Ronde palen verpakt (RP VERPAKT)"),
]

ALL_TAB_LABELS = dict(STANDARD_TABS_META + OPTIONAL_TABS_META)

KNOWN_FRENCH_CLIENTS = {
    "ARDENNES CLOTURES",
    "CLÔTURES OLIVIER",
    "CLOTURES OLIVIER",
    "CODEN TRADE SPRL",
    "EASYFENCE SRL",
    "ESPACE CHASSART",
    "EURO-FENCING",
    "LEBOUTTE & CIE SPRL",
    "MATTELE ALAIN",
    "BOLOGNE",
    "MOULAN",
}

INTERNAL_COLS_TO_STRIP = {
    "RP": {"K", "L", "M"},
    "DS 2D": {"H", "I", "J", "K", "L"},
    "PREST 3D": {"H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"},
    "VPI": {"H", "I", "J", "K", "L"},
    "VPZ": {"H", "I", "J"},
    "VPUS": {"K", "L", "M"},
    "VPUB": {"H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"},
    "SKP": {"H", "I", "J", "K", "L"},
    "MEHA GP": {"J", "K", "L", "M", "N"},
    "EP": {"K", "L"},
    "DP": {"K", "L"},
    "EPDS": {"I", "J", "K"},
    "DPDS": {"I", "J", "K"},
    "MAATP": {"F", "G", "H", "I", "J", "K"},
    "PK": {"L", "M", "N", "O", "P", "Q", "R", "S", "T"},
    "PK DS": {"L", "M", "N", "O", "P", "Q", "R", "S"},
    "RP SCHETS": {"H", "I", "J", "K", "L", "M", "N", "O"},
    "RP PRES": {"H", "I", "J", "K", "L", "M", "N", "O", "P"},
    "VPB": {"H", "I", "J", "K", "L", "M", "N"},
    "RP VERPAKT": {"J", "K", "L", "M", "N", "O", "P"},
}


# ==============================================================================
# CAMPAGNE INSTELLINGEN & VARIABELE DATUMS
# ==============================================================================

DEFAULT_SETTINGS = {
    "ingangsdatum_iso": "2026-06-01",
    "bron_nl": "Art. Prijzen_BRON_20260601.xlsx",
    "bron_fr": "Art. Prijzen_BRON_20260601.xlsx",
    "template_nl": "prijsstijging_algemeen_nl.md",
    "template_fr": "prijsstijging_algemeen_fr.md",
    "pas_begindatum_aan_in_excel": False,
}


def load_campaign_settings():
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(SETTINGS_JSON_PATH):
        save_campaign_settings(DEFAULT_SETTINGS)
        return dict(DEFAULT_SETTINGS)
    try:
        with open(SETTINGS_JSON_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        merged = dict(DEFAULT_SETTINGS)
        merged.update(data)
        return merged
    except Exception:
        return dict(DEFAULT_SETTINGS)


def save_campaign_settings(settings):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(SETTINGS_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2, ensure_ascii=False)


def get_date_variables(ingangsdatum_iso="2026-06-01"):
    """Berekent alle datum-formaten (kort, lang NL, lang FR, bestandsnaam-code, Excel serial) vanuit YYYY-MM-DD."""
    try:
        dt = datetime.strptime(ingangsdatum_iso, "%Y-%m-%d")
    except ValueError:
        dt = datetime(2026, 6, 1)
    datum_code = dt.strftime("%Y%m%d")
    datum_kort = f"{dt.day}/{dt.month}/{str(dt.year)[2:]}"
    datum_volledig = dt.strftime("%d/%m/%Y")
    datum_lang_nl = f"{dt.day} {NL_MAANDEN[dt.month]} {dt.year}"
    datum_lang_fr = f"{'1er' if dt.day == 1 else dt.day} {FR_MAANDEN[dt.month]} {dt.year}"
    excel_serial = (dt - datetime(1899, 12, 30)).days
    return {
        "ingangsdatum_iso": dt.strftime("%Y-%m-%d"),
        "datum_code": datum_code,
        "ingangsdatum_kort": datum_kort,
        "ingangsdatum_ddmmyyyy": datum_volledig,
        "ingangsdatum_lang_nl": datum_lang_nl,
        "ingangsdatum_lang_fr": datum_lang_fr,
        "excel_serial": excel_serial,
    }


def format_smart_greeting(customer):
    """
    Zorgt dat elke mail een persoonlijke aanspreking heeft:
    1. Als er een specifieke aanspreking is ingevuld (bv. 'Goedemiddag Anja,\\nGoedemiddag Mario'), gebruik die.
    2. Als er een contactpersoon (voornaam) is ingevuld, gebruik 'Goedemiddag [Voornaam],' / 'Bonjour [Voornaam],'.
    3. Anders: gebruik een nette bedrijfsaanspreking 'Beste partner bij [Klantnaam],' (nooit een kale 'Beste,').
    """
    lang = customer.get("taal", "NL")
    custom_aanspr = str(customer.get("aanspreking", "")).strip()
    contact = str(customer.get("contactpersoon", "")).strip()
    klantnaam = customer.get("klantnaam", "").replace(" (SPECIALS)", "").replace(" (STANDAARD_ROOD)", "")

    if custom_aanspr and custom_aanspr.lower() not in ("beste,", "bonjour,", "beste", "bonjour"):
        return custom_aanspr
    if contact:
        names = [n.strip() for n in re.split(r"[,&/]| en | et ", contact) if n.strip()]
        if len(names) > 1:
            prefix = "Bonjour" if lang == "FR" else "Goedemiddag"
            return "\n".join(f"{prefix} {n}," for n in names)
        return f"Bonjour {contact}," if lang == "FR" else f"Goedemiddag {contact},"

    if lang == "FR":
        return f"Bonjour l'équipe de {klantnaam},"
    return f"Beste team van {klantnaam},"


# ==============================================================================
# MARKDOWN TEMPLATES BEHEER & PARSER
# ==============================================================================

def list_mail_templates():
    """Geeft alle beschikbare Markdown mailtemplates uit templates/mails/ terug."""
    os.makedirs(TEMPLATES_DIR, exist_ok=True)
    templates = []
    for fn in sorted(os.listdir(TEMPLATES_DIR)):
        if not fn.endswith(".md"):
            continue
        p = os.path.join(TEMPLATES_DIR, fn)
        with open(p, "r", encoding="utf-8") as f:
            raw = f.read()
        meta, body = parse_markdown_template(raw)
        templates.append({
            "filename": fn,
            "naam": meta.get("naam", fn),
            "taal": meta.get("taal", "NL"),
            "onderwerp": meta.get("onderwerp", "Prijslijst {{ingangsdatum_kort}}"),
            "body": body,
            "raw": raw,
        })
    return templates


def parse_markdown_template(raw_text):
    """Splitst YAML-achtige frontmatter (--- ... ---) en de Markdown body."""
    meta = {}
    body = raw_text
    m = re.match(r"^---\s*\r?\n(.*?)\r?\n---\s*\r?\n(.*)$", raw_text, re.DOTALL)
    if m:
        fm_lines = m.group(1).splitlines()
        body = m.group(2).strip()
        for line in fm_lines:
            if ":" in line:
                k, v = line.split(":", 1)
                meta[k.strip().lower()] = v.strip()
    return meta, body


def save_mail_template(filename, naam, taal, onderwerp, body):
    os.makedirs(TEMPLATES_DIR, exist_ok=True)
    filename = (filename or "").strip()
    if not filename:
        # Genereer automatisch een nette bestandsnaam op basis van de templatenaam en taal
        slug = re.sub(r"[^a-z0-9]+", "_", naam.lower()).strip("_") or "template"
        filename = f"{slug}_{taal.lower()}.md"
    if not filename.endswith(".md"):
        filename += ".md"
    safe_fn = re.sub(r"[^a-zA-Z0-9_.-]", "_", filename)
    content = f"---\nnaam: {naam}\ntaal: {taal}\nonderwerp: {onderwerp}\n---\n{body.strip()}\n"
    path = os.path.join(TEMPLATES_DIR, safe_fn)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    return safe_fn


def markdown_to_html_and_text(md_text, lang="NL", cid_logo="vdm_logo_001"):
    """Converteert een ingevulde Markdown mailtemplate naar nette HTML (incl. VDM handtekening) en Plain Text."""
    if lang == "FR":
        sig_html = f"""<p style="color:#595959;">Cordialement,</p>
<p><strong>Nick Cuppers</strong><br>
Téléphone: +32 13 66 20 79<br>
Email: <a href="mailto:nick@vandermaesen-nv.com">nick@vandermaesen-nv.com</a><br>
<a href="https://vandermaesenafsluitingen.be/">https://vandermaesenafsluitingen.be/</a><br>
Hulsterweg 19 – BE 3980 Tessenderlo – Ham</p>
<p><img src="cid:{cid_logo}" alt="Vandermaesen" width="215"></p>"""
        sig_txt = (
            "Cordialement,\n\nNick Cuppers\nTéléphone: +32 13 66 20 79\n"
            "Email: nick@vandermaesen-nv.com\nhttps://vandermaesenafsluitingen.be/\n"
            "Hulsterweg 19 – BE 3980 Tessenderlo – Ham"
        )
    else:
        sig_html = f"""<p style="color:#595959;">Vriendelijke groeten,</p>
<p><strong>Nick Cuppers</strong><br>
Telefoon: +32 13 66 20 79<br>
Email: <a href="mailto:nick@vandermaesen-nv.com">nick@vandermaesen-nv.com</a><br>
<a href="https://vandermaesenafsluitingen.be/">https://vandermaesenafsluitingen.be/</a><br>
Hulsterweg 19 – BE 3980 Tessenderlo – Ham</p>
<p style="font-size:9.5pt;color:#333;"><strong>Wil u snel onze offerte ontvangen? Klik dan op onderstaande knop, vervolledig de stappen en u ontvangt spoedig onze aanbieding.</strong></p>
<p><a href="https://vandermaesenafsluitingen.be/" style="display:inline-block;background:#0078d4;color:#ffffff;text-decoration:none;padding:6px 16px;border-radius:999px;font-weight:bold;font-size:10pt;">Vraag een offerte aan</a></p>
<p><img src="cid:{cid_logo}" alt="Vandermaesen" width="215"></p>"""
        sig_txt = (
            "Vriendelijke groeten,\n\nNick Cuppers\nTelefoon: +32 13 66 20 79\n"
            "Email: nick@vandermaesen-nv.com\nhttps://vandermaesenafsluitingen.be/\n"
            "Hulsterweg 19 – BE 3980 Tessenderlo – Ham"
        )

    plain_text = md_text.replace("{{handtekening}}", sig_txt)
    plain_text = re.sub(r"\*\*(.*?)\*\*", r"\1", plain_text)
    plain_text = re.sub(r"`(.*?)`", r"\1", plain_text)

    # Bouw HTML blok per blok op
    blocks = re.split(r"\r?\n\r?\n+", md_text.strip())
    html_parts = []
    for block in blocks:
        b_strip = block.strip()
        if not b_strip:
            continue
        if b_strip == "{{handtekening}}":
            html_parts.append(sig_html)
            continue

        lines = b_strip.splitlines()
        if all(l.strip().startswith("- ") or l.strip().startswith("* ") for l in lines):
            li_items = []
            for l in lines:
                item_txt = l.strip()[2:]
                item_txt = _inline_md_to_html(item_txt)
                li_items.append(f"  <li>{item_txt}</li>")
            html_parts.append("<ul>\n" + "\n".join(li_items) + "\n</ul>")
        else:
            p_html = "<br>\n".join(_inline_md_to_html(l) for l in lines)
            html_parts.append(f"<p>{p_html}</p>")

    full_html = (
        '<html><body style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000; line-height: 1.45;">\n'
        + "\n".join(html_parts)
        + "\n</body></html>"
    )
    return full_html, plain_text


def _inline_md_to_html(text):
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)
    text = re.sub(r"\*(.+?)\*", r"<em>\1</em>", text)
    text = re.sub(r"`(.+?)`", r'<code style="background:#f1f5f9;padding:1px 4px;border-radius:3px;">\1</code>', text)
    text = re.sub(r"\[(.+?)\]\((.+?)\)", r'<a href="\2">\1</a>', text)
    return text


def render_customer_mail_content(customer, xlsx_filename, settings=None, template_override=None):
    """Vult alle variabelen ({{aanspreking}}, {{ingangsdatum_lang}}, {{klantnaam}}, etc.) in voor een klant."""
    if settings is None:
        settings = load_campaign_settings()
    dvars = get_date_variables(settings.get("ingangsdatum_iso", "2026-06-01"))
    lang = customer.get("taal", "NL")

    if template_override:
        meta, body = parse_markdown_template(template_override)
    else:
        tpl_fn = settings.get("template_fr" if lang == "FR" else "template_nl", "")
        tpl_path = os.path.join(TEMPLATES_DIR, tpl_fn)
        if not os.path.exists(tpl_path):
            tpl_path = os.path.join(
                TEMPLATES_DIR,
                "prijsstijging_algemeen_fr.md" if lang == "FR" else "prijsstijging_algemeen_nl.md"
            )
        with open(tpl_path, "r", encoding="utf-8") as f:
            meta, body = parse_markdown_template(f.read())

    clean_klantnaam = customer["klantnaam"].replace(" (SPECIALS)", "").replace(" (STANDAARD_ROOD)", "")
    aanspreking = format_smart_greeting(customer)
    ingangsdatum_lang = dvars["ingangsdatum_lang_fr"] if lang == "FR" else dvars["ingangsdatum_lang_nl"]

    replacements = {
        "{{aanspreking}}": aanspreking,
        "{{contactpersoon}}": customer.get("contactpersoon") or clean_klantnaam,
        "{{klantnaam}}": clean_klantnaam,
        "{{klantnr}}": customer.get("klantnr") or "",
        "{{ingangsdatum_lang}}": ingangsdatum_lang,
        "{{ingangsdatum_kort}}": dvars["ingangsdatum_kort"],
        "{{ingangsdatum_ddmmyyyy}}": dvars["ingangsdatum_ddmmyyyy"],
        "{{datum_code}}": dvars["datum_code"],
        "{{bijlage_naam}}": xlsx_filename,
    }

    subj = meta.get("onderwerp", "Prijsstijging {{ingangsdatum_kort}}")
    filled_body = body
    for k, v in replacements.items():
        subj = subj.replace(k, str(v))
        filled_body = filled_body.replace(k, str(v))

    html_body, plain_body = markdown_to_html_and_text(filled_body, lang=lang)
    return subj, html_body, plain_body


# ==============================================================================
# EXCEL HULPFUNCTIES & PARSERS
# ==============================================================================

def split_cell_ref(ref):
    m = re.match(r"^([A-Z]+)(\d+)$", ref)
    if not m:
        return "", 0
    return m.group(1), int(m.group(2))


def col_to_idx(col_str):
    idx = 0
    for ch in col_str:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def idx_to_col(idx):
    res = []
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        res.append(chr(ord("A") + rem))
    return "".join(reversed(res))


def excel_serial_to_date_str(val):
    try:
        serial = int(float(val))
        if 40000 <= serial <= 60000:
            dt = datetime(1899, 12, 30) + timedelta(days=serial)
            return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return str(val)


def parse_rp_discount(raw_val):
    if raw_val is None or str(raw_val).strip() == "":
        return 0.0, ""
    s = str(raw_val).strip()
    m = re.search(r"([+-]?\d+(?:[.,]\d+)?)\s*%", s)
    if m:
        val = float(m.group(1).replace(",", ".")) / 100.0
        return round(val, 6), s
    try:
        num = float(s.replace(",", "."))
        if abs(num) < 0.5 and num != 0:
            return round(num, 6), f"{num*100:.2f}% (Blauwe lijst)"
        return round(num / 100.0, 6), ""
    except ValueError:
        return 0.0, s


def parse_band_discount(raw_val):
    if raw_val is None or str(raw_val).strip() == "":
        return {"stuk_pct": 0.0, "pallet_pct": 0.05, "op_aanvraag": False, "raw": ""}
    s = str(raw_val).strip()
    if "aanvr" in s.lower():
        return {"stuk_pct": 0.0, "pallet_pct": 0.0, "op_aanvraag": True, "raw": s}
    m_both = re.search(r"\(1\)\s*(\d+(?:[.,]\d+)?)\s*-\s*\(VP\)\s*(\d+(?:[.,]\d+)?)", s, re.I)
    if m_both:
        return {
            "stuk_pct": round(float(m_both.group(1).replace(",", ".")) / 100.0, 6),
            "pallet_pct": round(float(m_both.group(2).replace(",", ".")) / 100.0, 6),
            "op_aanvraag": False,
            "raw": s,
        }
    m_vp = re.search(r"\(VP\)\s*(\d+(?:[.,]\d+)?)", s, re.I)
    if m_vp:
        return {
            "stuk_pct": 0.0,
            "pallet_pct": round(float(m_vp.group(1).replace(",", ".")) / 100.0, 6),
            "op_aanvraag": False,
            "raw": s,
        }
    try:
        v = float(s.replace(",", "."))
        pct = v / 100.0 if v >= 1 else v
        return {"stuk_pct": 0.0, "pallet_pct": round(pct, 6), "op_aanvraag": False, "raw": s}
    except ValueError:
        return {"stuk_pct": 0.0, "pallet_pct": 0.05, "op_aanvraag": False, "raw": s}


def parse_ep_dp_discount(raw_val):
    if raw_val is None or str(raw_val).strip() == "":
        return 0.0, 0.0, ""
    s = str(raw_val).strip()
    m = re.search(r"\(1\)\s*(\d+(?:[.,]\d+)?)\s*-\s*\(10\)\s*(\d+(?:[.,]\d+)?)", s)
    if m:
        d1 = round(float(m.group(1).replace(",", ".")) / 100.0, 6)
        d10 = round(float(m.group(2).replace(",", ".")) / 100.0, 6)
        return d1, d10, s
    m_single = re.search(r"(\d+(?:[.,]\d+)?)", s)
    if m_single:
        d = round(float(m_single.group(1).replace(",", ".")) / 100.0, 6)
        return d, d, s
    return 0.0, 0.0, s


def determine_tabs_for_customer(row_dict, group_name):
    q_note = str(row_dict.get("Q", "")).strip()
    q_upper = q_note.upper()

    if group_name == "SPECIALS" and q_note:
        if "ENKEL POORTKADERS + VDM-BAND" in q_upper:
            return ["VDM BAND", "PK", "PK DS"]
        if "ENKEL POORTEN + VDM-BAND" in q_upper:
            return ["VDM BAND", "EP", "DP", "EPDS", "DPDS", "MAATP"]
        if "ENKEL POORTEN" in q_upper:
            tabs = ["EP", "DP", "EPDS", "DPDS", "MAATP"]
            if str(row_dict.get("N", "")).strip().upper() == "X":
                tabs += ["PK", "PK DS", "PP SETS", "SLOTEN"]
            return tabs
        if "STEENKORFPALEN" in q_upper:
            return ["SKP"]
        if "TOEBEHOREN, GP, BAND, POORTEN" in q_upper:
            return ["TOEBEH", "VDM GP", "VLP GP", "MEHA GP", "VDM BAND", "EP", "DP", "EPDS", "DPDS", "MAATP"]
        if q_upper == "VPZ, GP, BAND":
            return ["VPZ", "VDM GP", "VLP GP", "MEHA GP", "VDM BAND"]
        if "VPI, VPZ, VPUS, VPUB" in q_upper:
            return ["VPI", "VPZ", "VPUS", "VPUB", "VP + VOET", "SKP", "VDM GP", "VLP GP", "MEHA GP"]
        if "GP, VPUB, BAND" in q_upper:
            return ["VPUB", "VDM GP", "VLP GP", "MEHA GP", "VDM BAND"]

    tabs = [t[0] for t in STANDARD_TABS_META]
    if "PRESTIPLI NETTEN NVT" in q_upper:
        tabs = [t for t in tabs if t not in ("DS 2D", "PREST 3D")]

    if str(row_dict.get("N", "")).strip().upper() == "X":
        tabs.extend(["PK", "PK DS", "PP SETS", "SLOTEN"])
    if str(row_dict.get("I", "")).strip().upper() == "X":
        tabs.append("RP SCHETS")
    if str(row_dict.get("D", "")).strip().upper() == "X":
        tabs.append("RP PRES")
    if str(row_dict.get("H", "")).strip().upper() == "X":
        tabs.append("VPB")
    if str(row_dict.get("E", "")).strip().upper() == "X":
        tabs.append("RP VERPAKT")

    return tabs


def parse_voorwaarden_excel(xlsx_path=VOORWAARDEN_PATH):
    ns = {"main": NS_MAIN}
    with zipfile.ZipFile(xlsx_path, "r") as z:
        ss = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss_xml = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in ss_xml.findall(".//main:si", ns):
                ss.append("".join(t.text or "" for t in si.findall(".//main:t", ns)))
        sh_xml = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))

        customers = []
        current_group = "STANDAARD_ROOD"

        for r in sh_xml.findall(".//main:row", ns):
            r_idx = int(r.attrib.get("r"))
            if r_idx <= 3:
                continue
            row_dict = {}
            for c in r.findall("main:c", ns):
                ref = c.attrib.get("r")
                col, _ = split_cell_ref(ref)
                ctype = c.attrib.get("t")
                v_el = c.find("main:v", ns)
                val = v_el.text if v_el is not None else ""
                if ctype == "s" and val.isdigit():
                    val = ss[int(val)]
                if val is not None and str(val).strip() != "":
                    row_dict[col] = str(val).strip()

            b_val = row_dict.get("B", "")
            if not b_val:
                continue

            if "OUDE CATEGORIE BLAUW" in b_val.upper():
                current_group = "CATEGORIE_BLAUW"
                continue
            elif b_val.upper() == "SPECIALS":
                current_group = "SPECIALS"
                continue
            elif "LANGE TIJD INACTIEF" in b_val.upper():
                current_group = "INACTIEF"
                continue
            elif "SPORADISCH AFHALEN" in b_val.upper():
                current_group = "SPORADISCH"
                continue
            elif "NIET MEER ACTIEF" in b_val.upper():
                current_group = "STOPGEZET"
                continue

            klantnr = row_dict.get("A", "")
            klantnaam = b_val.strip()
            if any(c["klantnaam"] == klantnaam for c in customers):
                klantnaam = f"{klantnaam} ({current_group})"

            rp_disc, rp_note = parse_rp_discount(row_dict.get("C"))
            band_info = parse_band_discount(row_dict.get("L"))
            ep_1, ep_10, _ = parse_ep_dp_discount(row_dict.get("O"))

            bph_raw = row_dict.get("M", "")
            bph_pct = 0.05 if "5" in bph_raw else 0.0

            op_raw = row_dict.get("P", "5")
            try:
                op_pct = float(op_raw.replace(",", ".")) / 100.0
            except ValueError:
                op_pct = 0.05

            q_note = row_dict.get("Q", "")
            geen_meerkost_7016 = "GEEN MEERKOST VOOR RAL 7016" in q_note.upper()

            meha_mode = "STANDAARD"
            if str(row_dict.get("K", "")).strip().upper() == "X":
                if "LENAERS" in klantnaam.upper():
                    meha_mode = "LENAERS"
                elif "DE PRIL" in klantnaam.upper() or "LM" in klantnaam.upper():
                    meha_mode = "LMDP"

            lang = "FR" if klantnaam.upper() in KNOWN_FRENCH_CLIENTS or any(
                w in klantnaam.upper() for w in ["SPRL", "SRL", "CLÔTURES", "CLOTURES"]
            ) else "NL"

            is_anma = "ANMABETON" in klantnaam.upper()
            email_addr = "anmabeton@hotmail.be" if is_anma else ""
            contactpersoon = "Anja en Mario" if is_anma else ""
            greeting = (
                "Goedemiddag Anja,\nGoedemiddag Mario"
                if is_anma
                else ""
            )

            tabs = determine_tabs_for_customer(row_dict, current_group)
            active = current_group in ("STANDAARD_ROOD", "CATEGORIE_BLAUW", "SPECIALS")

            cust_obj = {
                "id": f"R{r_idx:03d}",
                "row_nr": r_idx,
                "actief": active,
                "groep": current_group,
                "klantnr": klantnr,
                "klantnaam": klantnaam,
                "taal": lang,
                "contactpersoon": contactpersoon,
                "email": email_addr,
                "aanspreking": greeting,
                "korting_rp": rp_disc,
                "korting_rp_note": rp_note,
                "korting_vpi": round(float(row_dict.get("G", "0").replace(",", ".")) / 100.0, 6)
                if row_dict.get("G", "").replace(",", ".").replace("-", "").isdigit()
                else 0.0,
                "korting_band_stuk": band_info["stuk_pct"],
                "korting_band_pallet": band_info["pallet_pct"],
                "band_op_aanvraag": band_info["op_aanvraag"],
                "korting_bph_pallet": bph_pct,
                "korting_ep_lt10": ep_1,
                "korting_ep_ge10": ep_10,
                "toeslag_op": op_pct,
                "toebeh_palletprijzen": str(row_dict.get("F", "")).strip().upper() == "X",
                "special_vpu": str(row_dict.get("J", "")).strip().upper() == "X",
                "meha_mode": meha_mode,
                "pk_apart": str(row_dict.get("N", "")).strip().upper() == "X",
                "special_pres": str(row_dict.get("D", "")).strip().upper() == "X",
                "special_verpakt": str(row_dict.get("E", "")).strip().upper() == "X",
                "special_vpb": str(row_dict.get("H", "")).strip().upper() == "X",
                "special_schets": str(row_dict.get("I", "")).strip().upper() == "X",
                "geen_meerkost_7016": geen_meerkost_7016,
                "opmerking": q_note,
                "tabbladen": tabs,
            }
            if not cust_obj["aanspreking"]:
                cust_obj["aanspreking"] = format_smart_greeting(cust_obj)
            customers.append(cust_obj)

        return customers


def save_config_json(customers, path=CONFIG_JSON_PATH):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(customers, f, indent=2, ensure_ascii=False)


def read_master_klantenconfig_excel(xlsx_path=MASTER_EXCEL_PATH):
    """
    Leest wijzigingen die Nick rechtstreeks in Excel in 00_MASTER_KLANTENCONFIG.xlsx heeft gedaan
    (zoals de kolom 'Aanspreking Mail', 'E-mailadres', 'Contactpersoon' of kortingen) terug in!
    """
    if not os.path.exists(xlsx_path):
        return None
    ns = {"main": NS_MAIN}
    with zipfile.ZipFile(xlsx_path, "r") as z:
        ss = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss_xml = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in ss_xml.findall(".//main:si", ns):
                ss.append("".join(t.text or "" for t in si.findall(".//main:t", ns)))
        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_xml}
        first_sh = wb_xml.find(".//main:sheet", ns)
        if first_sh is None:
            return None
        rid = first_sh.attrib.get(f"{{{NS_REL}}}id")
        target = "xl/" + rid_to_target[rid].lstrip("/")
        sh_xml = ET.fromstring(z.read(target))

        rows_by_id = {}
        headers = {}
        for r in sh_xml.findall(".//main:row", ns):
            r_idx = int(r.attrib.get("r", "0"))
            row_vals = {}
            for c in r.findall("main:c", ns):
                col, _ = split_cell_ref(c.attrib.get("r", ""))
                ctype = c.attrib.get("t")
                v_el = c.find("main:v", ns)
                is_t = c.find(".//main:t", ns)
                if ctype == "s" and v_el is not None and v_el.text and v_el.text.isdigit():
                    val = ss[int(v_el.text)]
                elif is_t is not None and is_t.text is not None:
                    val = is_t.text
                elif v_el is not None and v_el.text is not None:
                    val = v_el.text
                else:
                    val = ""
                row_vals[col] = str(val).strip()
            if r_idx == 1:
                headers = {v: k for k, v in row_vals.items() if v}
            else:
                cid = row_vals.get(headers.get("ID", "A"), "")
                if cid:
                    rows_by_id[cid] = (headers, row_vals)
        return rows_by_id


def parse_master_klantenconfig_complete(xlsx_path=MASTER_EXCEL_PATH):
    """
    Reconstrueert de volledige klantenlijst rechtstreeks vanuit 00_MASTER_KLANTENCONFIG.xlsx.
    Werkt ook op een verse installatie waar data/klanten_config.json nog niet bestaat!
    """
    rows_by_id = read_master_klantenconfig_excel(xlsx_path)
    if not rows_by_id:
        return []
    customers = []
    for idx, (cid, (hdr, rv)) in enumerate(rows_by_id.items(), start=4):
        def g(col_name, default=""):
            col_let = hdr.get(col_name)
            return rv.get(col_let, default) if col_let else default

        def gf(col_name, default=0.0):
            try:
                return round(float(g(col_name, str(default)).replace(",", ".")) / 100.0, 6)
            except ValueError:
                return default

        tabs_raw = g("Actieve Tabbladen", "")
        tabs = [t.strip() for t in tabs_raw.split(",") if t.strip()]
        if not tabs:
            tabs = [t[0] for t in STANDARD_TABS_META]

        aanspr_raw = g("Aanspreking Mail", "")
        aanspr = "\n".join(p.strip() for p in aanspr_raw.split("/") if p.strip())

        m = re.search(r"\d+", cid)
        row_nr = int(m.group(0)) if m else idx

        cust = {
            "id": cid,
            "row_nr": row_nr,
            "actief": g("Actief", "JA").upper() == "JA",
            "groep": g("Groep", "STANDAARD_ROOD"),
            "klantnr": g("Klantnr (Navision)", ""),
            "klantnaam": g("Klantnaam", ""),
            "taal": g("Taal", "NL").upper() or "NL",
            "contactpersoon": g("Contactpersoon (Voornaam)", ""),
            "email": g("E-mailadres", ""),
            "aanspreking": aanspr,
            "korting_rp": gf("% Korting RP/SP", 0.0),
            "korting_rp_note": "",
            "korting_vpi": gf("% Korting VPI", 0.0),
            "korting_band_stuk": gf("% Korting Band (1st)", 0.0),
            "korting_band_pallet": gf("% Korting Band (Pallet)", 0.05),
            "band_op_aanvraag": g("Band Op Aanvraag", "NEE").upper() == "JA",
            "korting_bph_pallet": gf("% Korting BPH (Pallet)", 0.05),
            "korting_ep_lt10": gf("% Korting Poorten (<10)", 0.0),
            "korting_ep_ge10": gf("% Korting Poorten (>=10)", 0.0),
            "toeslag_op": gf("% Toeslag Los (OP)", 0.05),
            "toebeh_palletprijzen": g("Toebeh. Palletprijzen", "NEE").upper() == "JA",
            "special_vpu": g("Speciale VPU Tabel", "NEE").upper() == "JA",
            "meha_mode": g("MEHA Prijskolom", "STANDAARD") or "STANDAARD",
            "pk_apart": g("PK/PP/Sloten Apart", "NEE").upper() == "JA",
            "special_pres": g("Speciale RP PRES", "NEE").upper() == "JA",
            "special_verpakt": g("Speciale RP VERPAKT", "NEE").upper() == "JA",
            "special_vpb": g("Speciale VPB", "NEE").upper() == "JA",
            "special_schets": g("Speciale RP SCHETS", "NEE").upper() == "JA",
            "geen_meerkost_7016": g("Geen Meerkost RAL 7016", "NEE").upper() == "JA",
            "opmerking": g("Speciale Opmerkingen", ""),
            "tabbladen": tabs,
        }
        if not cust["aanspreking"]:
            cust["aanspreking"] = format_smart_greeting(cust)
        customers.append(cust)
    return customers


def get_files_status():
    """Controleert of de 2 verplichte gevoelige Excel-bestanden aanwezig zijn."""
    settings = load_campaign_settings()
    os.makedirs(RAW_DIR, exist_ok=True)
    bron_nl_name = settings.get("bron_nl", "Art. Prijzen_BRON_20260601.xlsx")
    bron_fr_name = settings.get("bron_fr", "")
    bron_nl_path = os.path.join(RAW_DIR, bron_nl_name)

    # Als de exacte bestandsnaam niet bestaat maar er staat wel een andere *BRON*.xlsx in _raw, pak die
    if not os.path.exists(bron_nl_path):
        raw_files = sorted(f for f in os.listdir(RAW_DIR) if f.endswith(".xlsx") and "BRON" in f.upper() and "FRANS" not in f.upper())
        if raw_files:
            bron_nl_name = raw_files[0]
            bron_nl_path = os.path.join(RAW_DIR, bron_nl_name)
            settings["bron_nl"] = bron_nl_name
            save_campaign_settings(settings)

    raw_fr_files = sorted(f for f in os.listdir(RAW_DIR) if f.endswith(".xlsx") and "FRANS" in f.upper())
    if raw_fr_files and (not bron_fr_name or not os.path.exists(os.path.join(RAW_DIR, bron_fr_name))):
        bron_fr_name = raw_fr_files[0]
        settings["bron_fr"] = bron_fr_name
        save_campaign_settings(settings)

    has_bron_nl = os.path.exists(bron_nl_path)
    has_bron_fr = bool(bron_fr_name and os.path.exists(os.path.join(RAW_DIR, bron_fr_name)) and "FRANS" in bron_fr_name.upper())
    has_master_config = os.path.exists(MASTER_EXCEL_PATH) or os.path.exists(CONFIG_JSON_PATH)

    return {
        "has_bron_nl": has_bron_nl,
        "bron_nl_name": bron_nl_name if has_bron_nl else None,
        "has_bron_fr": has_bron_fr,
        "bron_fr_name": bron_fr_name if has_bron_fr else None,
        "has_master_config": has_master_config,
        "master_config_name": os.path.basename(MASTER_EXCEL_PATH) if os.path.exists(MASTER_EXCEL_PATH) else None,
        "ready": has_bron_nl and has_master_config,
    }


def load_config_json(path=CONFIG_JSON_PATH):
    if not os.path.exists(path):
        if os.path.exists(MASTER_EXCEL_PATH):
            customers = parse_master_klantenconfig_complete(MASTER_EXCEL_PATH)
            save_config_json(customers, path)
            return customers
        elif os.path.exists(VOORWAARDEN_PATH):
            customers = parse_voorwaarden_excel(VOORWAARDEN_PATH)
            save_config_json(customers, path)
            generate_master_klantenconfig_excel(customers, MASTER_EXCEL_PATH)
            return customers
        else:
            return []
    with open(path, "r", encoding="utf-8") as f:
        customers = json.load(f)

    changed = False
    if os.path.exists(MASTER_EXCEL_PATH) and os.path.getmtime(MASTER_EXCEL_PATH) > os.path.getmtime(path) + 1:
        try:
            customers = parse_master_klantenconfig_complete(MASTER_EXCEL_PATH)
            changed = True
        except Exception:
            pass

    for c in customers:
        if "contactpersoon" not in c:
            c["contactpersoon"] = "Anja en Mario" if "ANMABETON" in c["klantnaam"].upper() else ""
            changed = True
        if c.get("aanspreking", "").strip().lower() in ("", "beste,", "bonjour,"):
            c["aanspreking"] = format_smart_greeting(c)
            changed = True
    if changed and customers:
        save_config_json(customers, path)
    return customers


def build_clean_xlsx_from_rows(sheets_data, out_path):
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        ct_parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
            '<Default Extension="xml" ContentType="application/xml"/>',
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        ]
        for idx in range(1, len(sheets_data) + 1):
            ct_parts.append(
                f'<Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            )
        ct_parts.append("</Types>")
        z.writestr("[Content_Types].xml", "".join(ct_parts))

        z.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>",
        )

        wb_sheets = []
        wb_rels = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        ]
        for idx, (sh_name, _) in enumerate(sheets_data, start=1):
            safe_name = sh_name.replace("&", "&amp;").replace('"', "&quot;")
            wb_sheets.append(f'<sheet name="{safe_name}" sheetId="{idx}" r:id="rId{idx}"/>')
            wb_rels.append(
                f'<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>'
            )
        wb_rels.append(
            f'<Relationship Id="rId{len(sheets_data)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        )
        wb_rels.append("</Relationships>")
        z.writestr("xl/_rels/workbook.xml.rels", "".join(wb_rels))

        z.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<sheets>{"".join(wb_sheets)}</sheets></workbook>',
        )

        styles_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="2">'
            '<font><sz val="10"/><name val="Calibri"/></font>'
            '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
            "</fonts>"
            '<fills count="3">'
            '<fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill>'
            '<fill><patternFill patternType="solid"><fgColor rgb="FF1B5E20"/><bgColor indexed="64"/></patternFill></fill>'
            "</fills>"
            '<borders count="2">'
            "<border><left/><right/><top/><bottom/><diagonal/></border>"
            '<border><left style="thin"><color rgb="FFD0D7DE"/></left><right style="thin"><color rgb="FFD0D7DE"/></right>'
            '<top style="thin"><color rgb="FFD0D7DE"/></top><bottom style="thin"><color rgb="FFD0D7DE"/></bottom><diagonal/></border>'
            "</borders>"
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="2">'
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>'
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
            "</cellXfs>"
            "</styleSheet>"
        )
        z.writestr("xl/styles.xml", styles_xml)

        for idx, (sh_name, rows) in enumerate(sheets_data, start=1):
            row_xml_parts = []
            max_cols = max((len(r) for r in rows), default=1)
            cols_xml = f'<cols><col min="1" max="{max_cols}" width="18" customWidth="1"/></cols>'
            for r_i, row_vals in enumerate(rows, start=1):
                style_attr = ' s="1"' if r_i == 1 else ' s="0"'
                c_parts = []
                for c_i, val in enumerate(row_vals, start=1):
                    col_letter = idx_to_col(c_i)
                    ref = f"{col_letter}{r_i}"
                    if val is None or val == "":
                        c_parts.append(f'<c r="{ref}"{style_attr}/>')
                    elif isinstance(val, (int, float)) and not isinstance(val, bool):
                        c_parts.append(f'<c r="{ref}"{style_attr}><v>{val}</v></c>')
                    else:
                        txt = (
                            str(val)
                            .replace("&", "&amp;")
                            .replace("<", "&lt;")
                            .replace(">", "&gt;")
                        )
                        c_parts.append(
                            f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{txt}</t></is></c>'
                        )
                row_xml_parts.append(f'<row r="{r_i}">{"".join(c_parts)}</row>')
            last_col = idx_to_col(max_cols)
            auto_filter = f'<autoFilter ref="A1:{last_col}{max(1, len(rows))}"/>'
            ws_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
                f"{cols_xml}<sheetData>{''.join(row_xml_parts)}</sheetData>{auto_filter}</worksheet>"
            )
            z.writestr(f"xl/worksheets/sheet{idx}.xml", ws_xml)


def generate_master_klantenconfig_excel(customers=None, out_path=MASTER_EXCEL_PATH):
    if customers is None:
        customers = load_config_json()

    headers = [
        "ID",
        "Actief",
        "Groep",
        "Klantnr (Navision)",
        "Klantnaam",
        "Taal",
        "Contactpersoon (Voornaam)",
        "E-mailadres",
        "Aanspreking Mail",
        "% Korting RP/SP",
        "% Korting VPI",
        "% Korting Band (1st)",
        "% Korting Band (Pallet)",
        "Band Op Aanvraag",
        "% Korting BPH (Pallet)",
        "% Korting Poorten (<10)",
        "% Korting Poorten (>=10)",
        "% Toeslag Los (OP)",
        "Toebeh. Palletprijzen",
        "Speciale VPU Tabel",
        "MEHA Prijskolom",
        "PK/PP/Sloten Apart",
        "Speciale RP PRES",
        "Speciale RP VERPAKT",
        "Speciale VPB",
        "Speciale RP SCHETS",
        "Geen Meerkost RAL 7016",
        "Aantal Tabbladen",
        "Actieve Tabbladen",
        "Speciale Opmerkingen",
    ]
    rows = [headers]
    for c in customers:
        rows.append([
            c["id"],
            "JA" if c["actief"] else "NEE",
            c["groep"],
            c["klantnr"],
            c["klantnaam"],
            c["taal"],
            c.get("contactpersoon", ""),
            c["email"],
            c["aanspreking"].replace("\n", " / "),
            round(c["korting_rp"] * 100, 2),
            round(c["korting_vpi"] * 100, 2),
            round(c["korting_band_stuk"] * 100, 2),
            round(c["korting_band_pallet"] * 100, 2),
            "JA" if c["band_op_aanvraag"] else "NEE",
            round(c["korting_bph_pallet"] * 100, 2),
            round(c["korting_ep_lt10"] * 100, 2),
            round(c["korting_ep_ge10"] * 100, 2),
            round(c["toeslag_op"] * 100, 2),
            "JA" if c["toebeh_palletprijzen"] else "NEE",
            "JA" if c["special_vpu"] else "NEE",
            c["meha_mode"],
            "JA" if c["pk_apart"] else "NEE",
            "JA" if c["special_pres"] else "NEE",
            "JA" if c["special_verpakt"] else "NEE",
            "JA" if c["special_vpb"] else "NEE",
            "JA" if c["special_schets"] else "NEE",
            "JA" if c["geen_meerkost_7016"] else "NEE",
            len(c["tabbladen"]),
            ", ".join(c["tabbladen"]),
            c["opmerking"],
        ])

    build_clean_xlsx_from_rows([("Klantenconfiguratie", rows)], out_path)
    return out_path


# ==============================================================================
# KERN-ENGINE: GENEREER 100% GEZUIVERDE KLANT-EXCEL VANUIT DE BRONFILE
# ==============================================================================

class SharedStringsManager:
    def __init__(self, xml_bytes):
        self.ns = {"main": NS_MAIN}
        self.root = ET.fromstring(xml_bytes)
        self.strings = []
        self.str_to_idx = {}
        for idx, si in enumerate(self.root.findall("main:si", self.ns)):
            txt = "".join(t.text or "" for t in si.findall(".//main:t", self.ns))
            self.strings.append(txt)
            if txt not in self.str_to_idx:
                self.str_to_idx[txt] = idx

    def get(self, idx):
        if 0 <= idx < len(self.strings):
            return self.strings[idx]
        return ""

    def get_or_add(self, text):
        text = str(text)
        if text in self.str_to_idx:
            return self.str_to_idx[text]
        idx = len(self.strings)
        self.strings.append(text)
        self.str_to_idx[text] = idx
        si = ET.SubElement(self.root, f"{{{NS_MAIN}}}si")
        t = ET.SubElement(si, f"{{{NS_MAIN}}}t")
        if text.startswith(" ") or text.endswith(" "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
        return idx

    def to_bytes(self):
        self.root.set("count", str(len(self.strings)))
        self.root.set("uniqueCount", str(len(self.strings)))
        return ET.tostring(self.root, encoding="utf-8", xml_declaration=True)


def get_cell_val(c_el, ssm):
    if c_el is None:
        return None
    ctype = c_el.attrib.get("t")
    v_el = c_el.find(f"{{{NS_MAIN}}}v")
    if v_el is None or v_el.text is None:
        return None
    if ctype == "s" and v_el.text.isdigit():
        return ssm.get(int(v_el.text))
    return v_el.text


def get_cell_float(c_el, ssm):
    val = get_cell_val(c_el, ssm)
    if val is None:
        return None
    try:
        return float(val)
    except ValueError:
        return None


def set_cell_number(c_el, num_val):
    if "t" in c_el.attrib:
        del c_el.attrib["t"]
    f_el = c_el.find(f"{{{NS_MAIN}}}f")
    if f_el is not None:
        c_el.remove(f_el)
    v_el = c_el.find(f"{{{NS_MAIN}}}v")
    if v_el is None:
        v_el = ET.SubElement(c_el, f"{{{NS_MAIN}}}v")
    v_el.text = repr(float(num_val)) if isinstance(num_val, float) else str(num_val)


def set_cell_string(c_el, text_val, ssm):
    f_el = c_el.find(f"{{{NS_MAIN}}}f")
    if f_el is not None:
        c_el.remove(f_el)
    c_el.set("t", "s")
    v_el = c_el.find(f"{{{NS_MAIN}}}v")
    if v_el is None:
        v_el = ET.SubElement(c_el, f"{{{NS_MAIN}}}v")
    v_el.text = str(ssm.get_or_add(text_val))


def build_row_map(sh_root):
    rmap = {}
    for r_el in sh_root.findall(f".//{{{NS_MAIN}}}row"):
        r_idx = int(r_el.attrib.get("r", "0"))
        cmap = {}
        for c_el in r_el.findall(f"{{{NS_MAIN}}}c"):
            col, _ = split_cell_ref(c_el.attrib.get("r", ""))
            if col:
                cmap[col] = c_el
        rmap[r_idx] = (r_el, cmap)
    return rmap


def apply_customer_to_sheet(sh_name, sh_root, customer, ssm, rp_netto_map, settings=None):
    rmap = build_row_map(sh_root)
    klantnaam = customer["klantnaam"].replace(" (SPECIALS)", "").replace(" (STANDAARD_ROOD)", "")
    dvars = get_date_variables((settings or {}).get("ingangsdatum_iso", "2026-06-01"))

    if sh_name == "Welkom":
        if 6 in rmap and "C" in rmap[6][1]:
            set_cell_string(rmap[6][1]["C"], klantnaam, ssm)
        if 7 in rmap and "I" in rmap[7][1]:
            set_cell_string(rmap[7][1]["I"], f"Vanaf {dvars['ingangsdatum_ddmmyyyy']}", ssm)
        active_tabs = [t for t in customer["tabbladen"] if t != "Welkom"]
        sheet_data = sh_root.find(f"{{{NS_MAIN}}}sheetData")
        style_a = rmap.get(9, (None, {}))[1].get("A").attrib.get("s") if 9 in rmap and "A" in rmap[9][1] else None
        style_b = rmap.get(9, (None, {}))[1].get("B").attrib.get("s") if 9 in rmap and "B" in rmap[9][1] else None
        for r_idx in list(rmap.keys()):
            if r_idx >= 9:
                sheet_data.remove(rmap[r_idx][0])
        for idx, t_name in enumerate(active_tabs, start=1):
            r_num = 7 + idx * 2
            r_el = ET.SubElement(sheet_data, f"{{{NS_MAIN}}}row", {"r": str(r_num), "spans": "1:9"})
            c_a = ET.SubElement(r_el, f"{{{NS_MAIN}}}c", {"r": f"A{r_num}"})
            if style_a:
                c_a.set("s", style_a)
            set_cell_string(c_a, f"Tabblad {idx}", ssm)
            c_b = ET.SubElement(r_el, f"{{{NS_MAIN}}}c", {"r": f"B{r_num}"})
            if style_b:
                c_b.set("s", style_b)
            set_cell_string(c_b, ALL_TAB_LABELS.get(t_name, t_name), ssm)
        return

    if 5 in rmap:
        if "C" in rmap[5][1] and get_cell_val(rmap[5][1]["C"], ssm) in ("BRON ROOD", klantnaam):
            set_cell_string(rmap[5][1]["C"], klantnaam, ssm)
        elif "B" in rmap[5][1] and get_cell_val(rmap[5][1]["B"], ssm) in ("BRON ROOD", klantnaam):
            set_cell_string(rmap[5][1]["B"], klantnaam, ssm)

    if sh_name == "RP":
        disc = customer["korting_rp"]
        if 7 in rmap and "I" in rmap[7][1]:
            set_cell_number(rmap[7][1]["I"], disc)
        for r_idx, (_, cmap) in rmap.items():
            if r_idx >= 9 and "H" in cmap and "I" in cmap:
                bruto = get_cell_float(cmap["H"], ssm)
                if bruto is not None:
                    netto = bruto - (bruto * disc)
                    set_cell_number(cmap["I"], netto)
                    rp_netto_map[r_idx] = netto

    elif sh_name == "SP":
        disc = customer["korting_rp"]
        if 7 in rmap and "G" in rmap[7][1]:
            set_cell_number(rmap[7][1]["G"], disc)
        for r_idx, (_, cmap) in rmap.items():
            if r_idx >= 9 and "F" in cmap and "G" in cmap:
                bruto = get_cell_float(cmap["F"], ssm)
                if bruto is not None:
                    set_cell_number(cmap["G"], bruto - (bruto * disc))

    elif sh_name == "RP PRES":
        for r_idx in range(9, 22):
            if r_idx in rmap and "F" in rmap[r_idx][1]:
                rp_row = r_idx + 6
                if rp_row in rp_netto_map:
                    set_cell_number(rmap[r_idx][1]["F"], rp_netto_map[rp_row] + 0.33)

    elif sh_name == "TOEBEH":
        if not customer["toebeh_palletprijzen"]:
            if 7 in rmap and "J" in rmap[7][1]:
                rmap[7][0].remove(rmap[7][1]["J"])
                del rmap[7][1]["J"]
            for r_idx, (r_el, cmap) in rmap.items():
                for col_del in ("I", "J"):
                    if col_del in cmap:
                        r_el.remove(cmap[col_del])
                        del cmap[col_del]
                if "K" in cmap:
                    c_k = cmap["K"]
                    c_k.set("r", f"I{r_idx}")
                    cmap["I"] = c_k
                    del cmap["K"]
                for col_extra in ("L", "M", "N"):
                    if col_extra in cmap:
                        r_el.remove(cmap[col_extra])
                        del cmap[col_extra]
            for mc in sh_root.findall(f".//{{{NS_MAIN}}}mergeCell"):
                ref = mc.attrib.get("ref", "")
                if ref.endswith("K203") or ref.endswith("K199") or ref.endswith("K200"):
                    mc.set("ref", ref.replace(":K", ":I"))

    elif sh_name in ("VPI", "VPZ", "VPUS"):
        price_col = "I" if sh_name == "VPUS" else "F"
        vpi_disc = customer["korting_vpi"] if sh_name == "VPI" else 0.0
        sorted_rows = sorted(r for r in rmap.keys() if r >= 8)
        for r_idx in sorted_rows:
            _, cmap = rmap[r_idx]
            cfg = str(get_cell_val(cmap.get("B"), ssm) or "").strip().upper()
            if customer["geen_meerkost_7016"] and cfg == "RAL 7016":
                art_nr = get_cell_val(cmap.get("A"), ssm)
                if (r_idx + 1) in rmap:
                    next_cmap = rmap[r_idx + 1][1]
                    if get_cell_val(next_cmap.get("A"), ssm) == art_nr and price_col in next_cmap:
                        p_6005 = get_cell_float(next_cmap[price_col], ssm)
                        if p_6005 is not None and price_col in cmap:
                            set_cell_number(cmap[price_col], p_6005)
            if vpi_disc != 0.0 and price_col in cmap:
                p = get_cell_float(cmap[price_col], ssm)
                if p is not None:
                    set_cell_number(cmap[price_col], p * (1.0 - vpi_disc))

    elif sh_name == "VPUB":
        for r_idx in range(12, 55):
            if r_idx not in rmap:
                continue
            _, cmap = rmap[r_idx]
            if customer["special_vpu"] and "S" in cmap and "F" in cmap:
                spec_p = get_cell_float(cmap["S"], ssm)
                if spec_p is not None:
                    set_cell_number(cmap["F"], spec_p)
        if customer["geen_meerkost_7016"]:
            for r_idx in range(12, 55):
                if r_idx in rmap and (r_idx + 1) in rmap:
                    _, cmap = rmap[r_idx]
                    _, next_cmap = rmap[r_idx + 1]
                    cfg = str(get_cell_val(cmap.get("B"), ssm) or "").strip().upper()
                    if cfg == "RAL 7016" and get_cell_val(cmap.get("A"), ssm) == get_cell_val(next_cmap.get("A"), ssm):
                        p_6005 = get_cell_float(next_cmap.get("F"), ssm)
                        if p_6005 is not None and "F" in cmap:
                            set_cell_number(cmap["F"], p_6005)

    elif sh_name == "MEHA GP":
        mode = customer["meha_mode"]
        if mode in ("LENAERS", "LMDP"):
            src_col = "M" if mode == "LENAERS" else "N"
            for r_idx in range(10, 41):
                if r_idx in rmap:
                    _, cmap = rmap[r_idx]
                    spec_p = get_cell_float(cmap.get(src_col), ssm)
                    if spec_p is not None:
                        for target_col in ("F", "G", "H"):
                            if target_col in cmap:
                                set_cell_number(cmap[target_col], spec_p)
        sheet_data = sh_root.find(f"{{{NS_MAIN}}}sheetData")
        for r_idx in list(rmap.keys()):
            if r_idx > 43:
                sheet_data.remove(rmap[r_idx][0])
                del rmap[r_idx]

    elif sh_name == "VDM BAND":
        stuk_disc = customer["korting_band_stuk"]
        pal_disc = customer["korting_band_pallet"]
        if 8 in rmap and "H" in rmap[8][1]:
            if customer["band_op_aanvraag"]:
                set_cell_string(rmap[8][1]["H"], "Op aanvraag", ssm)
            else:
                set_cell_number(rmap[8][1]["H"], pal_disc)
        for r_idx in range(10, 23):
            if r_idx not in rmap:
                continue
            _, cmap = rmap[r_idx]
            if "F" in cmap:
                base_f = get_cell_float(cmap["F"], ssm)
                if base_f is not None:
                    if customer["band_op_aanvraag"]:
                        set_cell_string(cmap["F"], "Op aanvraag", ssm)
                        if "H" in cmap:
                            set_cell_string(cmap["H"], "Op aanvraag", ssm)
                    else:
                        if stuk_disc > 0:
                            set_cell_number(cmap["F"], base_f * (1.0 - stuk_disc))
                        if "H" in cmap:
                            set_cell_number(cmap["H"], base_f * (1.0 - pal_disc))

    elif sh_name == "BPH":
        bph_disc = customer["korting_bph_pallet"]
        op_toeslag = customer["toeslag_op"]
        for r_idx in range(9, 20):
            if r_idx not in rmap:
                continue
            _, cmap = rmap[r_idx]
            doos_h = get_cell_float(cmap.get("H"), ssm)
            if doos_h is not None:
                if "F" in cmap:
                    set_cell_number(cmap["F"], doos_h * (1.0 + op_toeslag))
                if "J" in cmap:
                    set_cell_number(cmap["J"], doos_h * (1.0 - bph_disc))

    elif sh_name in ("EP", "DP"):
        d1 = customer["korting_ep_lt10"]
        d10 = customer["korting_ep_ge10"]
        if 9 in rmap:
            if "F" in rmap[9][1]:
                set_cell_number(rmap[9][1]["F"], d1)
            if "H" in rmap[9][1]:
                set_cell_number(rmap[9][1]["H"], d10)
        for r_idx, (_, cmap) in rmap.items():
            if r_idx >= 11 and "E" in cmap:
                bruto = get_cell_float(cmap["E"], ssm)
                if bruto is not None:
                    if "G" in cmap:
                        set_cell_number(cmap["G"], bruto * (1.0 - d1))
                    if "I" in cmap:
                        set_cell_number(cmap["I"], bruto * (1.0 - d10))

    elif sh_name in ("EPDS", "DPDS"):
        d1 = customer["korting_ep_lt10"]
        d10 = customer["korting_ep_ge10"]
        disc_row = 10 if sh_name == "EPDS" else 9
        start_row = 12 if sh_name == "EPDS" else 11
        if disc_row in rmap:
            if "F" in rmap[disc_row][1]:
                set_cell_number(rmap[disc_row][1]["F"], d1)
            if "G" in rmap[disc_row][1]:
                set_cell_number(rmap[disc_row][1]["G"], d10)
        for r_idx, (_, cmap) in rmap.items():
            if r_idx >= start_row and "E" in cmap:
                bruto = get_cell_float(cmap["E"], ssm)
                if bruto is not None:
                    if "F" in cmap:
                        set_cell_number(cmap["F"], bruto * (1.0 - d1))
                    if "G" in cmap:
                        set_cell_number(cmap["G"], bruto * (1.0 - d10))

    elif sh_name in ("PK", "PK DS", "SLOTEN"):
        d1 = customer["korting_ep_lt10"]
        d10 = customer["korting_ep_ge10"]
        if 7 in rmap:
            if "G" in rmap[7][1]:
                set_cell_number(rmap[7][1]["G"], d1)
            if "I" in rmap[7][1]:
                set_cell_number(rmap[7][1]["I"], d10)
        if sh_name in ("PK", "PK DS") and 3 in rmap:
            sh_root.find(f"{{{NS_MAIN}}}sheetData").remove(rmap[3][0])
            del rmap[3]
        for r_idx, (_, cmap) in rmap.items():
            if r_idx >= 9 and "F" in cmap:
                bruto = get_cell_float(cmap["F"], ssm)
                if bruto is not None:
                    if "H" in cmap:
                        set_cell_number(cmap["H"], bruto * (1.0 - d1))
                    if "J" in cmap:
                        set_cell_number(cmap["J"], bruto * (1.0 - d10))

    elif sh_name == "PP SETS":
        d1 = customer["korting_ep_lt10"]
        d10 = customer["korting_ep_ge10"]
        if 7 in rmap:
            if "H" in rmap[7][1]:
                set_cell_number(rmap[7][1]["H"], d1)
            if "J" in rmap[7][1]:
                set_cell_number(rmap[7][1]["J"], d10)
        for r_idx, (_, cmap) in rmap.items():
            if r_idx >= 9 and "G" in cmap:
                bruto = get_cell_float(cmap["G"], ssm)
                if bruto is not None:
                    if "I" in cmap:
                        set_cell_number(cmap["I"], bruto * (1.0 - d1))
                    if "K" in cmap:
                        set_cell_number(cmap["K"], bruto * (1.0 - d10))

    cols_to_strip = INTERNAL_COLS_TO_STRIP.get(sh_name, set())
    if cols_to_strip:
        for r_idx, (r_el, cmap) in rmap.items():
            for c_del in list(cols_to_strip):
                if c_del in cmap:
                    r_el.remove(cmap[c_del])
                    del cmap[c_del]
        merge_parent = sh_root.find(f"{{{NS_MAIN}}}mergeCells")
        if merge_parent is not None:
            for mc in list(merge_parent.findall(f"{{{NS_MAIN}}}mergeCell")):
                ref = mc.attrib.get("ref", "")
                parts = ref.split(":")
                if any(split_cell_ref(p)[0] in cols_to_strip for p in parts):
                    merge_parent.remove(mc)
            if len(merge_parent.findall(f"{{{NS_MAIN}}}mergeCell")) == 0:
                sh_root.remove(merge_parent)

    cols_el = sh_root.find(f"{{{NS_MAIN}}}cols")
    if cols_el is not None:
        for col_el in list(cols_el.findall(f"{{{NS_MAIN}}}col")):
            if col_el.attrib.get("hidden") in ("1", "true"):
                cols_el.remove(col_el)

    for f_parent in sh_root.findall(f".//{{{NS_MAIN}}}c"):
        f_el = f_parent.find(f"{{{NS_MAIN}}}f")
        if f_el is not None:
            f_parent.remove(f_el)


def generate_customer_workbook(customer, bron_path=None, out_dir=None, settings=None):
    if settings is None:
        settings = load_campaign_settings()
    dvars = get_date_variables(settings.get("ingangsdatum_iso", "2026-06-01"))

    if bron_path is None:
        lang_file = settings.get("bron_fr" if customer.get("taal") == "FR" else "bron_nl", "Art. Prijzen_BRON_20260601.xlsx")
        candidate = os.path.join(RAW_DIR, lang_file)
        bron_path = candidate if os.path.exists(candidate) else BRON_NL_PATH

    if out_dir is None:
        out_dir = os.path.join(OUT_DIR, "01_Prijslijsten_Klanten")
    os.makedirs(out_dir, exist_ok=True)

    safe_klant = re.sub(r'[\\/*?:"<>|]', "_", customer["klantnaam"]).strip()
    lang_tag = "_FRANS" if customer.get("taal") == "FR" else ""
    out_filename = f"Art. Prijzen_{safe_klant}{lang_tag}_{dvars['datum_code']}.xlsx"
    out_path = os.path.join(out_dir, out_filename)

    allowed_sheets = ["Welkom"] + [t for t in customer["tabbladen"] if t != "Welkom"]

    with zipfile.ZipFile(bron_path, "r") as zin:
        ssm = SharedStringsManager(zin.read("xl/sharedStrings.xml"))
        wb_root = ET.fromstring(zin.read("xl/workbook.xml"))
        rels_root = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
        ct_root = ET.fromstring(zin.read("[Content_Types].xml"))

        for ac in list(wb_root.findall("{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent")):
            wb_root.remove(ac)
        for ext_refs in list(wb_root.findall(f"{{{NS_MAIN}}}externalReferences")):
            wb_root.remove(ext_refs)

        rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root}

        sheets_el = wb_root.find(f"{{{NS_MAIN}}}sheets")
        kept_sheet_targets = set()
        kept_rids = set()
        old_idx_to_new_idx = {}
        kept_sheet_names = []

        all_sheet_els = list(sheets_el.findall(f"{{{NS_MAIN}}}sheet"))
        for old_idx, sh_el in enumerate(all_sheet_els):
            name = sh_el.attrib.get("name")
            rid = sh_el.attrib.get(f"{{{NS_REL}}}id")
            if name in allowed_sheets:
                new_idx = len(kept_sheet_names)
                old_idx_to_new_idx[old_idx] = new_idx
                kept_sheet_names.append(name)
                kept_rids.add(rid)
                kept_sheet_targets.add("xl/" + rid_to_target[rid].lstrip("/"))
            else:
                sheets_el.remove(sh_el)

        for bv in wb_root.findall(f".//{{{NS_MAIN}}}workbookView"):
            bv.set("activeTab", "0")

        def_names_el = wb_root.find(f"{{{NS_MAIN}}}definedNames")
        if def_names_el is not None:
            for dn in list(def_names_el.findall(f"{{{NS_MAIN}}}definedName")):
                lsid = dn.attrib.get("localSheetId")
                if lsid is not None and lsid.isdigit():
                    old_i = int(lsid)
                    if old_i in old_idx_to_new_idx:
                        dn.set("localSheetId", str(old_idx_to_new_idx[old_i]))
                        if dn.text and "TOEBEH!" in dn.text and not customer["toebeh_palletprijzen"]:
                            dn.text = dn.text.replace("$K$", "$I$")
                    else:
                        def_names_el.remove(dn)

        for rel in list(rels_root):
            rid = rel.attrib.get("Id")
            rtype = rel.attrib.get("Type", "")
            if "externalLink" in rtype or "calcChain" in rtype:
                rels_root.remove(rel)
            elif "worksheet" in rtype and rid not in kept_rids:
                rels_root.remove(rel)

        for ov in list(ct_root.findall(f"{{{NS_CT}}}Override")):
            pn = ov.attrib.get("PartName", "").lstrip("/")
            if "externalLink" in pn or "calcChain" in pn:
                ct_root.remove(ov)
            elif pn.startswith("xl/worksheets/sheet") and pn not in kept_sheet_targets:
                ct_root.remove(ov)

        modified_sheets_bytes = {}
        rp_netto_map = {}

        sheet_order = sorted(
            [(sh_el.attrib.get("name"), "xl/" + rid_to_target[sh_el.attrib.get(f"{{{NS_REL}}}id")].lstrip("/"))
             for sh_el in sheets_el.findall(f"{{{NS_MAIN}}}sheet")],
            key=lambda x: (0 if x[0] == "RP" else 1),
        )

        for sh_name, target_path in sheet_order:
            sh_root = ET.fromstring(zin.read(target_path))
            apply_customer_to_sheet(sh_name, sh_root, customer, ssm, rp_netto_map, settings=settings)
            modified_sheets_bytes[target_path] = ET.tostring(sh_root, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                fn = item.filename
                if fn.startswith("xl/externalLinks/") or fn == "xl/calcChain.xml":
                    continue
                if fn.startswith("xl/worksheets/sheet") and fn not in kept_sheet_targets:
                    continue
                if fn.startswith("xl/worksheets/_rels/sheet"):
                    base_sheet = fn.replace("xl/worksheets/_rels/", "xl/worksheets/").replace(".rels", "")
                    if base_sheet not in kept_sheet_targets:
                        continue
                if fn == "xl/workbook.xml":
                    zout.writestr(fn, ET.tostring(wb_root, encoding="utf-8", xml_declaration=True))
                elif fn == "xl/_rels/workbook.xml.rels":
                    zout.writestr(fn, ET.tostring(rels_root, encoding="utf-8", xml_declaration=True))
                elif fn == "[Content_Types].xml":
                    zout.writestr(fn, ET.tostring(ct_root, encoding="utf-8", xml_declaration=True))
                elif fn == "xl/sharedStrings.xml":
                    zout.writestr(fn, ssm.to_bytes())
                elif fn in modified_sheets_bytes:
                    zout.writestr(fn, modified_sheets_bytes[fn])
                else:
                    zout.writestr(fn, zin.read(fn))

    return out_path


# ==============================================================================
# E-MAIL GENERATOR (.EML MET MARKDOWN TEMPLATE, LOGO & BIJLAGE)
# ==============================================================================

_CACHED_LOGO_BYTES = None


def get_vdm_logo_bytes():
    global _CACHED_LOGO_BYTES
    if _CACHED_LOGO_BYTES is not None:
        return _CACHED_LOGO_BYTES
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            _CACHED_LOGO_BYTES = f.read()
        return _CACHED_LOGO_BYTES
    _CACHED_LOGO_BYTES = b""
    return _CACHED_LOGO_BYTES


def generate_customer_eml(customer, xlsx_path, out_dir=None, settings=None):
    """Genereert een .eml bestand op basis van de gekozen Markdown template + variabele datums + slimme aanspreking."""
    if settings is None:
        settings = load_campaign_settings()
    dvars = get_date_variables(settings.get("ingangsdatum_iso", "2026-06-01"))

    if out_dir is None:
        out_dir = os.path.join(OUT_DIR, "02_Outlook_Conceptmails")
    os.makedirs(out_dir, exist_ok=True)

    safe_klant = re.sub(r'[\\/*?:"<>|]', "_", customer["klantnaam"]).strip()
    eml_path = os.path.join(out_dir, f"Mail_{safe_klant}_{dvars['datum_code']}.eml")

    xlsx_filename = os.path.basename(xlsx_path)
    subj, html_body, plain_body = render_customer_mail_content(customer, xlsx_filename, settings=settings)

    clean_klantnaam = customer["klantnaam"].replace(" (SPECIALS)", "").replace(" (STANDAARD_ROOD)", "")
    msg = EmailMessage()
    msg["From"] = "Nick Cuppers <Nick@vandermaesen-nv.com>"
    to_email = customer.get("email") or f"info@{re.sub(r'[^a-z0-9]', '', clean_klantnaam.lower())}.be"
    msg["To"] = f"{clean_klantnaam} <{to_email}>"
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = subj
    msg["X-Unsent"] = "1"

    msg.set_content(plain_body)
    msg.add_alternative(html_body, subtype="html")

    logo_bytes = get_vdm_logo_bytes()
    if logo_bytes:
        msg.get_payload()[1].add_related(
            logo_bytes, maintype="image", subtype="png", cid="<vdm_logo_001>", filename="image001.png"
        )

    with open(xlsx_path, "rb") as f:
        xlsx_data = f.read()
    msg.add_attachment(
        xlsx_data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=xlsx_filename,
    )

    with open(eml_path, "wb") as f:
        f.write(bytes(msg))

    return eml_path


# ==============================================================================
# NAVISION EXPORT & AUTO-UPDATER
# ==============================================================================

def generate_navision_export(customers=None, bron_path=BRON_NL_PATH, out_dir=None, settings=None):
    if settings is None:
        settings = load_campaign_settings()
    dvars = get_date_variables(settings.get("ingangsdatum_iso", "2026-06-01"))
    bdat_nav = dvars["ingangsdatum_ddmmyyyy"]

    if customers is None:
        customers = [c for c in load_config_json() if c.get("actief")]
    if out_dir is None:
        out_dir = os.path.join(OUT_DIR, "03_Navision_Export")
    os.makedirs(out_dir, exist_ok=True)

    rgk_rows = [[
        "Verkooptype",
        "Verkoopcode (Klantnr)",
        "Klantnaam",
        "Type",
        "Code (Artikelkortingsgroep)",
        "Minimumaantal",
        "Regelkorting %",
        "Begindatum",
        "Opmerking",
    ]]
    for c in customers:
        knr = c["klantnr"] or "ONTBREEKT"
        knaam = c["klantnaam"]
        if c["korting_rp"] != 0:
            for grp in ("RP", "SP"):
                if grp in c["tabbladen"]:
                    rgk_rows.append(["Klant", knr, knaam, "Artikelkortingsgroep", grp, 1, round(c["korting_rp"] * 100, 2), bdat_nav, c["korting_rp_note"]])
        if c["korting_vpi"] != 0 and "VPI" in c["tabbladen"]:
            rgk_rows.append(["Klant", knr, knaam, "Artikelkortingsgroep", "VPI", 1, round(c["korting_vpi"] * 100, 2), bdat_nav, ""])
        if "VDM BAND" in c["tabbladen"] and not c["band_op_aanvraag"]:
            if c["korting_band_stuk"] > 0:
                rgk_rows.append(["Klant", knr, knaam, "Artikelkortingsgroep", "VDMBAND", 1, round(c["korting_band_stuk"] * 100, 2), bdat_nav, "Stukkorting"])
            if c["korting_band_pallet"] > 0:
                rgk_rows.append(["Klant", knr, knaam, "Artikelkortingsgroep", "VDMBAND", 200, round(c["korting_band_pallet"] * 100, 2), bdat_nav, "Volle pallet"])
        if "BPH" in c["tabbladen"] and c["korting_bph_pallet"] > 0:
            rgk_rows.append(["Klant", knr, knaam, "Artikelkortingsgroep", "BPH", 600, round(c["korting_bph_pallet"] * 100, 2), bdat_nav, "Volle pallet"])
        if c["korting_ep_lt10"] > 0 or c["korting_ep_ge10"] > 0:
            poort_groups = [("EP", "EP", 1, 10), ("DP", "DP", 1, 10), ("EPDS", "EPDS", 1, 10), ("DPDS", "DPDS", 1, 10)]
            if c["pk_apart"]:
                poort_groups += [("PK", "PK", 1, 10), ("PK DS", "PKDS", 1, 10), ("PP SETS", "PP", 1, 30), ("SLOTEN", "SLOT", 1, 30)]
            for tab_name, nav_code, min1, min2 in poort_groups:
                if tab_name in c["tabbladen"]:
                    rgk_rows.append(["Klant", knr, knaam, "Artikelkortingsgroep", nav_code, min1, round(c["korting_ep_lt10"] * 100, 2), bdat_nav, f"Afname < {min2}"])
                    rgk_rows.append(["Klant", knr, knaam, "Artikelkortingsgroep", nav_code, min2, round(c["korting_ep_ge10"] * 100, 2), bdat_nav, f"Afname >= {min2}"])

    art_rows = [[
        "Tabblad",
        "Artikelengroep",
        "Artikelnr.",
        "Variantcode (Configuratie)",
        "Omschrijving",
        "Eenheidscode (EC)",
        "Minimumaantal (Staffel)",
        "Eenheidsprijs (€/st)",
        "Begindatum",
    ]]
    with zipfile.ZipFile(bron_path, "r") as z:
        ssm = SharedStringsManager(z.read("xl/sharedStrings.xml"))
        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root}

        for sh_el in wb_root.findall(f".//{{{NS_MAIN}}}sheet"):
            sh_name = sh_el.attrib.get("name")
            if sh_name in ("Welkom", "MAATP", "RP VERPAKT", "VP + VOET"):
                continue
            target = "xl/" + rid_to_target[sh_el.attrib.get(f"{{{NS_REL}}}id")].lstrip("/")
            sh_root = ET.fromstring(z.read(target))
            rmap = build_row_map(sh_root)
            art_grp = get_cell_val(rmap.get(2, (None, {}))[1].get("A"), ssm) or sh_name

            for r_idx in sorted(rmap.keys()):
                if r_idx < 8:
                    continue
                _, cmap = rmap[r_idx]
                art_nr = str(get_cell_val(cmap.get("A"), ssm) or get_cell_val(cmap.get("B"), ssm) or "").strip()
                if not re.match(r"^[PE]\d{5,7}$", art_nr):
                    continue
                if sh_name in ("VDM GP", "VLP GP"):
                    cfg = str(get_cell_val(cmap.get("C"), ssm) or "")
                    oms = str(get_cell_val(cmap.get("D"), ssm) or "")
                    ec = str(get_cell_val(cmap.get("E"), ssm) or "ST")
                    per_pal = get_cell_float(cmap.get("F"), ssm) or 1
                    bdat = excel_serial_to_date_str(get_cell_val(cmap.get("J"), ssm) or "46174")
                    for col_p, mult in [("G", 1), ("H", 3), ("I", 5)]:
                        p = get_cell_float(cmap.get(col_p), ssm)
                        if p is not None:
                            art_rows.append([sh_name, art_grp, art_nr, cfg, oms, ec, int(per_pal * mult), round(p, 4), bdat])
                elif sh_name == "TOEBEH":
                    cfg = str(get_cell_val(cmap.get("B"), ssm) or "")
                    oms = str(get_cell_val(cmap.get("C"), ssm) or "")
                    ec = str(get_cell_val(cmap.get("D"), ssm) or "ST")
                    bdat = excel_serial_to_date_str(get_cell_val(cmap.get("K"), ssm) or "46174")
                    for q_col, p_col in [("E", "F"), ("G", "H"), ("I", "J")]:
                        q = get_cell_float(cmap.get(q_col), ssm)
                        p = get_cell_float(cmap.get(p_col), ssm)
                        if q is not None and p is not None:
                            art_rows.append([sh_name, art_grp, art_nr, cfg, oms, ec, int(q), round(p, 4), bdat])
                elif sh_name == "RP":
                    cfg = str(get_cell_val(cmap.get("B"), ssm) or "")
                    oms = str(get_cell_val(cmap.get("C"), ssm) or "")
                    ec = str(get_cell_val(cmap.get("D"), ssm) or "ST")
                    q = get_cell_float(cmap.get("G"), ssm) or 1
                    p = get_cell_float(cmap.get("H"), ssm)
                    bdat = excel_serial_to_date_str(get_cell_val(cmap.get("J"), ssm) or "46174")
                    if p is not None:
                        art_rows.append([sh_name, art_grp, art_nr, cfg, oms, ec, int(q), round(p, 4), bdat])
                elif sh_name == "VPUS":
                    cfg = str(get_cell_val(cmap.get("B"), ssm) or "")
                    oms = str(get_cell_val(cmap.get("C"), ssm) or "")
                    ec = str(get_cell_val(cmap.get("D"), ssm) or "ST")
                    q = get_cell_float(cmap.get("H"), ssm) or 1
                    p = get_cell_float(cmap.get("I"), ssm)
                    bdat = excel_serial_to_date_str(get_cell_val(cmap.get("J"), ssm) or "46174")
                    if p is not None:
                        art_rows.append([sh_name, art_grp, art_nr, cfg, oms, ec, int(q), round(p, 4), bdat])
                else:
                    cfg = str(get_cell_val(cmap.get("B"), ssm) or "")
                    oms = str(get_cell_val(cmap.get("C"), ssm) or "")
                    ec = str(get_cell_val(cmap.get("D"), ssm) or "ST")
                    q = get_cell_float(cmap.get("E"), ssm) or 1
                    p = get_cell_float(cmap.get("F"), ssm) or get_cell_float(cmap.get("E"), ssm)
                    bdat = excel_serial_to_date_str(get_cell_val(cmap.get("G"), ssm) or get_cell_val(cmap.get("H"), ssm) or "46174")
                    if p is not None:
                        art_rows.append([sh_name, art_grp, art_nr, cfg, oms, ec, int(q), round(p, 4), bdat])

    out_xlsx = os.path.join(out_dir, f"Navision_Prijzen_en_RGK_{dvars['datum_code']}.xlsx")
    build_clean_xlsx_from_rows(
        [
            ("1_Navision_RGK_Klantkorting", rgk_rows),
            ("2_Navision_Artikelprijzen", art_rows),
        ],
        out_xlsx,
    )
    return out_xlsx


def audit_workbook_security(xlsx_path):
    with zipfile.ZipFile(xlsx_path, "r") as z:
        ext_links = [f for f in z.namelist() if "externalLink" in f]
        wb_str = z.read("xl/workbook.xml").decode("utf-8", errors="ignore")
        has_abspath = "absPath" in wb_str or "V:\\" in wb_str
        total_formulas = 0
        hidden_cols = 0
        for fn in z.namelist():
            if fn.startswith("xl/worksheets/sheet") and fn.endswith(".xml"):
                sh_root = ET.fromstring(z.read(fn))
                total_formulas += len(sh_root.findall(f".//{{{NS_MAIN}}}f"))
                for c in sh_root.findall(f".//{{{NS_MAIN}}}col"):
                    if c.attrib.get("hidden") in ("1", "true"):
                        hidden_cols += 1
        return {
            "bestand": os.path.basename(xlsx_path),
            "externe_koppelingen": len(ext_links),
            "formules": total_formulas,
            "verborgen_kolommen": hidden_cols,
            "lokale_serverpaden": has_abspath,
            "veilig": (len(ext_links) == 0 and total_formulas == 0 and hidden_cols == 0 and not has_abspath),
        }


def check_and_apply_updates():
    """
    Haalt via puur Python (zonder dat Git op Nicks HP EliteBook nodig is)
    de nieuwste versie van vdm_engine.py en app.py op van GitHub.
    Raakt NOOIT klantdata (confidential/) aan!
    """
    updated_files = []
    ssl_ctx = ssl._create_unverified_context()
    for filename in ("vdm_engine.py", "app.py"):
        url = f"{GITHUB_RAW_BASE_URL.rstrip('/')}/{filename}"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "VDM-Studio-Updater"})
            with urllib.request.urlopen(req, timeout=2.0, context=ssl_ctx) as resp:
                if resp.status == 200:
                    new_bytes = resp.read()
                    if b"VDM Prijslijst Studio" in new_bytes:
                        local_path = os.path.join(BASE_DIR, filename)
                        with open(local_path, "rb") as f:
                            old_bytes = f.read()
                        if new_bytes != old_bytes:
                            with open(local_path, "wb") as f:
                                f.write(new_bytes)
                            updated_files.append(filename)
        except Exception as e:
            return {"ok": False, "updated": [], "message": f"Update-server nog niet bereikbaar of offline ({e})."}

    if updated_files:
        return {"ok": True, "updated": updated_files, "message": f"Geüpdatet: {', '.join(updated_files)}."}
    return {"ok": True, "updated": [], "message": f"Je werkt al met de nieuwste versie (v{APP_VERSION})!"}
