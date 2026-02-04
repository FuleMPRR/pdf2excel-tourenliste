import re
from io import BytesIO

import pandas as pd
import pdfplumber
import streamlit as st


# =============================
# Hilfsfunktionen
# =============================
ARTICLE_TOKENS = ["DGB 2023", "GB 2023", "KB"]


def normalize(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def extract_phones(text: str) -> str:
    phones = re.findall(r"(?:\+|00)?\d[\d\s]{7,}\d", text)
    phones = [normalize(p) for p in phones]
    seen, out = set(), []
    for p in phones:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return " / ".join(out)


def extract_contact_name(text: str) -> str:
    m = re.search(r"(?:\+|00)?\d[\d\s]{7,}\d", text)
    if not m:
        return ""
    return normalize(text[: m.start()].strip(" /"))


def extract_article(text: str) -> str:
    for a in ARTICLE_TOKENS:
        if a in text:
            return a
    return ""


def extract_plz_ort(text: str) -> str:
    m = re.search(r"\b(\d{4})\s+([A-Za-zÄÖÜäöü\- ]+)", text)
    if not m:
        return ""
    return normalize(f"{m.group(1)} {m.group(2)}")


def extract_street(text: str) -> str:
    m = re.search(
        r"([A-Za-zÄÖÜäöü\-]+(?:strasse|straße|weg|platz|gasse|ring|allee)\s+\d+\w?)",
        text,
        re.IGNORECASE,
    )
    if m:
        return normalize(m.group(1))
    return ""


# =============================
# Record-Erkennung
# =============================
def split_records(lines):
    end_re = re.compile(r"\d+/\d+\.\d+\s+\d+\s+\d+$")
    records, buf = [], []

    for line in lines:
        line = line.strip()
        if not line:
            continue
        if (
            "Tourenliste per:" in line
            or line.startswith("103_Tourenliste")
            or line.startswith("Firma Ansprech")
            or line.startswith("Tour ")
            or "Seite:" in line
        ):
            continue

        buf.append(line)
        if end_re.search(line):
            records.append(buf)
            buf = []

    return records


def parse_record(block):
    end_re = re.compile(r"(?P<pos>\d+/\d+\.\d+)\s+(?P<adr>\d+)\s+(?P<rh>\d+)$")

    full = normalize(" ".join(block))
    firma_raw = normalize(block[0])

    contact_line = ""
    for l in block:
        if re.search(r"(?:\+|00)?\d[\d\s]{7,}\d", l):
            contact_line = l
            break

    ansprech = extract_contact_name(contact_line)
    telefon = extract_phones(full)
    artikel = extract_article(full)
    plz_ort = extract_plz_ort(full)
    strasse = extract_street(full)

    pos, adr, rh = "", "", ""
    m = end_re.search(block[-1])
    if m:
        pos, adr, rh = m.group("pos"), m.group("adr"), m.group("rh")

    # 👉 Firma + Ansprechpartner ZUSAMMEN
    if ansprech:
        firma = normalize(f"{firma_raw} – {ansprech}")
    else:
        firma = firma_raw

    bemerkung = full
    if artikel:
        bemerkung = normalize(full.split(artikel, 1)[-1])
    if pos:
        bemerkung = re.sub(rf"{re.escape(pos)}\s+{adr}\s+{rh}$", "", bemerkung).strip()

    return {
        "Firma": firma,
        "Telefon": telefon,
        "Strasse": strasse,
        "PLZ / Ort": plz_ort,
        "Artikel": artikel,
        "Bemerkung": bemerkung,
        "Position Box": pos,
        "Adr.-Nr.": adr,
        "Rhythmus": rh,
    }


def parse_tourenliste(pdf_bytes: bytes) -> pd.DataFrame:
    lines = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for p in pdf.pages:
            lines.extend((p.extract_text() or "").split("\n"))

    records = split_records(lines)
    rows = [parse_record(r) for r in records]

    df = pd.DataFrame(rows)
    df = df[df["Position Box"].str.match(r"\d+/\d+\.\d+", na=False)]
    df.reset_index(drop=True, inplace=True)

    return df


def df_to_excel(df: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Tour")
    return out.getvalue()


# =============================
# Streamlit UI
# =============================
st.set_page_config(page_title="PDF → Excel (Tourenliste)", layout="centered")
st.title("PDF → Excel (Tourenliste)")
st.write("PDF hochladen → Excel herunterladen. Keine Installation notwendig.")

uploaded = st.file_uploader("PDF Datei", type=["pdf"])

if uploaded:
    with st.spinner("Verarbeite PDF …"):
        df = parse_tourenliste(uploaded.read())

    st.success(f"Gefundene Eintraege: {len(df)}")
    st.dataframe(df, use_container_width=True)

    st.download_button(
        "Excel herunterladen",
        df_to_excel(df),
        file_name=uploaded.name.replace(".pdf", ".xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
