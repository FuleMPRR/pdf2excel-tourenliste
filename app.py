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
    phones = re.findall(r"(?:\+|00)?\d[\d\s]{7,}\d", text)
    phones = [normalize_spaces(p) for p in phones]
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
    m = re.search(r"\b(\d{4})\s+([A-Za-zÄÖÜäöü\-]+(?:\s+[A-Za-zÄÖÜäöü\-]+)*)\b", text)
    if not m:
        return ""
    return normalize_spaces(f"{m.group(1)} {m.group(2)}")


def extract_street(text: str) -> str:
    plz_pos = None
    m_plz = re.search(r"\b\d{4}\b", text)
    if m_plz:
        plz_pos = m_plz.start()
    pre = text[:plz_pos].strip() if plz_pos else text

    pre = normalize_spaces(pre)

    m1 = re.search(
        r"\b([A-Za-zÄÖÜäöü\-]+(?:strasse|straße|weg|platz|gasse|ring|allee))\s+\d+\w?\b",
        pre,
        re.IGNORECASE,
    )
    if m1:
        return normalize_spaces(pre[m1.start():])

    m2 = re.search(r"([A-Za-zÄÖÜäöü\-]+\s+\d+\w?)\s*$", pre)
    if m2:
        return normalize_spaces(m2.group(1))

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
        if line.startswith("Tour "):
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

    full = normalize_spaces(" ".join(block_lines))

    pos_box = ""
    adr = ""
    rh = ""

    last = block_lines[-1]
    m_end = end_re.match(last)
    if m_end:
        pos_box = m_end.group("pos")
        adr = m_end.group("adr")
        rh = m_end.group("rh")

    # Firma ist in deiner Liste praktisch immer die erste Zeile
    firma_raw = normalize_spaces(block_lines[0])

    # Entferne reine Nummern-Zeilen (z.B. 8689, 8592) aus dem Block fuer Analyse
    cleaned_lines = [l for l in block_lines[1:] if not re.fullmatch(r"\d{2,6}", l.strip())]

    # Ansprechpartner / Telefon: erste Zeile mit Tel
    contact_line = ""
    for l in cleaned_lines:
        if re.search(r"(?:\+|00)?\d[\d\s]{7,}\d", l):
            contact_line = l
            break

    telefon = extract_phones(contact_line) if contact_line else extract_phones(full)
    ansprech = extract_contact_name(contact_line) if contact_line else ""

    # Artikel / PLZ / Strasse
    artikel = extract_article(full)
    plz_ort = split_plz_ort(full)

    street_candidate = ""
    for l in cleaned_lines:
        if re.search(r"(strasse|straße|weg|platz|gasse|ring|allee)\b", l, re.IGNORECASE):
            street_candidate = l
            break
    strasse = extract_street(street_candidate if street_candidate else full)

    # Bemerkung: alles nach Artikel (falls vorhanden), sonst fallback
    bemerkung = ""
    if artikel:
        idx = full.find(artikel)
        after = full[idx + len(artikel):].strip()
        if pos_box:
            after = re.sub(
                rf"\b{re.escape(pos_box)}\b\s+{re.escape(adr)}\s+{re.escape(rh)}\s*$",
                "",
                after,
            ).strip()
        bemerkung = after
    else:
        bemerkung = full

    bemerkung = normalize_spaces(bemerkung)

    # ✅ Wunsch: Ansprechperson mit Firma zusammennehmen
    if ansprech:
        firma = normalize_spaces(f"{firma_raw} - {ansprech}")
    else:
        firma = firma_raw

    return {
        "Firma": firma,
        "Ansprechperson": "",  # absichtlich leer, weil in Firma integriert
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
    all_lines = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            all_lines.extend(txt.split("\n"))

    blocks = parse_records_from_text(all_lines)
    rows = [parse_block(b) for b in blocks]

    df = pd.DataFrame(rows)

    df["Firma"] = df["Firma"].fillna("").map(normalize_spaces)
    df = df[df["Position Box"].str.match(r"^\d+/\d+\.\d+$", na=False)].reset_index(drop=True)

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
        st.warning("Keine Eintraege erkannt. Wenn du willst, schick einen Screenshot der Streamlit-Logs, dann passen wir den Parser noch enger an.")
