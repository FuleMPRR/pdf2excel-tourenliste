import re
from io import BytesIO

import pandas as pd
import pdfplumber
import streamlit as st


# -----------------------------
# Robust Parser: record-based
# -----------------------------
ARTICLE_TOKENS = ["DGB 2023", "GB 2023", "KB"]  # Reihenfolge wichtig (DGB/GB zuerst)


def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def extract_phones(text: str) -> str:
    """
    Extrahiert Telefonnummern robust aus einem Textblock.
    Akzeptiert +41..., 0041..., 071..., 079... etc.
    """
    # grob, aber brauchbar
    phones = re.findall(r"(?:\+|00)?\d[\d\s]{7,}\d", text)
    phones = [normalize_spaces(p) for p in phones]
    # Duplikate entfernen, Reihenfolge behalten
    seen = set()
    out = []
    for p in phones:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return " / ".join(out)


def extract_contact_name(text: str) -> str:
    """
    Nimmt den Teil vor der ersten erkannten Telefonnummer als Name.
    """
    # finde erste Tel-Position
    m = re.search(r"(?:\+|00)?\d[\d\s]{7,}\d", text)
    if not m:
        return ""
    name = text[: m.start()].strip(" /")
    return normalize_spaces(name)


def extract_article(text: str) -> str:
    for tok in ARTICLE_TOKENS:
        if tok in text:
            return tok
    return ""


def split_plz_ort(text: str) -> str:
    """
    Nimmt den ersten PLZ/Ort-Teil (4-stellige PLZ + Ort).
    Gibt z.B. '8583 Sulgen' zurueck.
    """
    # Beispiel: "... Kradolfstrasse 54 8583 Sulgen KB Werkstatt ..."
    m = re.search(r"\b(\d{4})\s+([A-Za-zÄÖÜäöü\-]+(?:\s+[A-Za-zÄÖÜäöü\-]+)*)\b", text)
    if not m:
        return ""
    return normalize_spaces(f"{m.group(1)} {m.group(2)}")


def extract_street(text: str) -> str:
    """
    Findet eine Strasse (typisch: '...strasse 54', 'Kirchplatz 5', 'Dorfstrasse 1a', 'Hintertschwil' etc.)
    Heuristik: nimm den Teil, der vor PLZ kommt und 'strasse/platz/weg' enthaelt ODER ein Wort + Hausnr.
    """
    # Teil vor PLZ (falls vorhanden)
    plz_pos = None
    m_plz = re.search(r"\b\d{4}\b", text)
    if m_plz:
        plz_pos = m_plz.start()
    pre = text[:plz_pos].strip() if plz_pos else text

    pre = normalize_spaces(pre)

    # Suche typische Muster
    m1 = re.search(r"\b([A-Za-zÄÖÜäöü\-]+(?:strasse|straße|weg|platz|gasse|ring|allee))\s+\d+\w?\b", pre, re.IGNORECASE)
    if m1:
        # nimm ab dem Start dieses Musters bis Ende von pre
        return normalize_spaces(pre[m1.start():])

    # Alternative: letztes "Wort + Hausnr" im pre
    m2 = re.search(r"([A-Za-zÄÖÜäöü\-]+\s+\d+\w?)\s*$", pre)
    if m2:
        return normalize_spaces(m2.group(1))

    # Falls z.B. nur Orts-/Weilername ohne Nummer (Hintertschwil)
    # nimm letzten Teil, wenn er nicht nach Tel aussieht
    if pre and len(pre.split()) <= 4:
        return pre

    return ""


def parse_records_from_text(lines):
    """
    Baut Record-Blöcke anhand des sicheren Endmarkers:
    <PositionBox> <AdrNr> <Rhyt> am Zeilenende.
    """
    end_re = re.compile(r"(?P<pos>\d+/\d+\.\d+)\s+(?P<adr>\d+)\s+(?P<rh>\d+)\s*$")

    blocks = []
    buf = []

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Kopf/Fuss raus
        if "Tourenliste per:" in line:
            continue
        if line.startswith("103_Tourenliste"):
            continue
        if line.startswith("Firma Ansprech"):
            continue
        if line.startswith("Tour ") and "Gesamt" in line:
            continue
        if line.startswith("Tour ") and "Gesamt" not in line:
            # z.B. "Tourenliste per: Woche 7 Tour: 86" (manchmal als Tour)
            continue
        if "Seite:" in line:
            continue

        buf.append(line)

        if end_re.search(line):
            blocks.append(buf)
            buf = []

    return blocks


def parse_block(block_lines):
    """
    Parse einen Record-Block zu Spalten.
    """
    end_re = re.compile(r"^(?P<pre>.*?)(?P<pos>\d+/\d+\.\d+)\s+(?P<adr>\d+)\s+(?P<rh>\d+)\s*$")

    # Record-Text
    full = normalize_spaces(" ".join(block_lines))

    # Position Box / Adr / Rhyt
    pos_box = ""
    adr = ""
    rh = ""

    # Suche die Zeile mit dem Endmarker (normalerweise letzte)
    last = block_lines[-1]
    m_end = end_re.match(last)
    if m_end:
        pos_box = m_end.group("pos")
        adr = m_end.group("adr")
        rh = m_end.group("rh")
        pre_last = normalize_spaces(m_end.group("pre"))
    else:
        pre_last = ""

    # Firma: in deiner Liste ist sie praktisch immer die erste Zeile
    firma = normalize_spaces(block_lines[0])

    # Entferne reine Nummern-Zeilen (z.B. 8689, 8592) aus dem Block fuer Analyse
    cleaned_lines = [l for l in block_lines[1:] if not re.fullmatch(r"\d{2,6}", l.strip())]

    # Ansprechpartner / Telefon: suche die erste Zeile mit einer Telefonnummer
    contact_line = ""
    for l in cleaned_lines:
        if re.search(r"(?:\+|00)?\d[\d\s]{7,}\d", l):
            contact_line = l
            break

    telefon = extract_phones(contact_line) if contact_line else extract_phones(full)
    ansprech = extract_contact_name(contact_line) if contact_line else ""

    # Artikel
    artikel = extract_article(full)

    # PLZ / Ort
    plz_ort = split_plz_ort(full)

    # Strasse
    # Oft steht Strasse in einer Zeile mit " / " (z.B. "... / Kradolfstrasse 54 8583 Sulgen ...")
    # Wir versuchen zuerst die Zeilen zu nehmen, die "strasse/platz/weg" enthalten.
    street_candidate = ""
    for l in cleaned_lines:
        if re.search(r"(strasse|straße|weg|platz|gasse|ring|allee)\b", l, re.IGNORECASE):
            street_candidate = l
            break
    strasse = extract_street(street_candidate if street_candidate else full)

    # Bemerkung:
    # Nimm alles nach Artikel (falls vorhanden) bis vor PositionBox (im letzten pre-Teil)
    bemerkung = ""

    if artikel:
        # split ab Artikel
        idx = full.find(artikel)
        after = full[idx + len(artikel):].strip()
        # PositionBox steht am Ende im last-line; entferne pos/adr/rh
        if pos_box:
            after = re.sub(rf"\b{re.escape(pos_box)}\b\s+{re.escape(adr)}\s+{re.escape(rh)}\s*$", "", after).strip()
        bemerkung = after
    else:
        # fallback: nimm pre_last (Teil vor PositionBox) minus bekannte Bestandteile
        bemerkung = pre_last

    bemerkung = normalize_spaces(bemerkung)

    return {
        "Firma": firma,
        "Ansprechperson": ansprech,
        "Telefon": telefon,
        "Strasse": strasse,
        "PLZ / Ort": plz_ort,
        "Artikel": artikel,
        "Bemerkung": bemerkung,
        "Position Box": pos_box,
        "Adr.-Nr.": adr,
        "Rhythmus": rh,
    }


def parse_tourenliste(pdf_bytes: bytes) -> pd.DataFrame:
    # Text aus allen Seiten ziehen
    all_lines = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            lines = txt.split("\n")
            all_lines.extend(lines)

    # Records bilden
    blocks = parse_records_from_text(all_lines)

    # Blocks parsen
    rows = [parse_block(b) for b in blocks]

    df = pd.DataFrame(rows)

    # Fix: Leere Firma raus / nur gültige PositionBox
    df["Firma"] = df["Firma"].fillna("").map(normalize_spaces)
    df = df[df["Position Box"].str.match(r"^\d+/\d+\.\d+$", na=False)].reset_index(drop=True)

    # Spaltenreihenfolge fix
    columns = [
        "Firma",
        "Ansprechperson",
        "Telefon",
        "Strasse",
        "PLZ / Ort",
        "Artikel",
        "Bemerkung",
        "Position Box",
        "Adr.-Nr.",
        "Rhythmus",
    ]
    for c in columns:
        if c not in df.columns:
            df[c] = ""
    df = df[columns]

    return df


def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Tour")
    return output.getvalue()


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="PDF → Excel (Tourenliste)", layout="centered")
st.title("PDF → Excel (Tourenliste)")
st.write("PDF hochladen, automatisch in Excel umwandeln und direkt herunterladen. Keine Installation auf dem Arbeits-PC noetig.")

uploaded = st.file_uploader("PDF Datei", type=["pdf"])

if uploaded:
    pdf_bytes = uploaded.read()

    with st.spinner("PDF wird verarbeitet ..."):
        df = parse_tourenliste(pdf_bytes)

    st.success(f"Gefundene Eintraege: {len(df)}")

    if len(df) > 0:
        st.dataframe(df, use_container_width=True)

        excel_bytes = df_to_xlsx_bytes(df)
        filename = uploaded.name.rsplit(".", 1)[0] + ".xlsx"

        st.download_button(
            label="Excel herunterladen",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning("Keine Eintraege erkannt. Falls du willst, poste einen Screenshot der Streamlit-Logs, dann machen wir den Parser noch enger passend.")
