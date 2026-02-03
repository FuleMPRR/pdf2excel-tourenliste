import re
from io import BytesIO

import pandas as pd
import pdfplumber
import streamlit as st


# -----------------------------
# PDF -> DataFrame Parser
# -----------------------------
def parse_tourenliste(pdf_bytes: bytes):
    rows = []
    current = {}

    # Regex-Patterns passend zu deiner Tourenliste
    pos_re = re.compile(r"\b\d+/\d+\.\d+\b")           # z.B. 86/12.0
    plz_re = re.compile(r"\b\d{4}\b")                  # Schweizer PLZ
    artikel_re = re.compile(r"\b(KB|GB 2023|DGB 2023)\b")

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = [l.strip() for l in text.split("\n") if l.strip()]

            for line in lines:
                # Kopf/Fuss ignorieren
                if (
                    "Tourenliste per:" in line
                    or line.startswith("103_Tourenliste")
                    or "Seite:" in line
                    or line.startswith("Tour ")
                    or line.startswith("Firma ")
                ):
                    continue

                # Neuer Datensatz = Firmenname
                if (
                    not pos_re.search(line)
                    and not line[:1].isdigit()
                    and len(line) > 2
                    and "Woche" not in line
                ):
                    if current.get("Position Box"):
                        rows.append(current)
                        current = {}

                    current["Firma"] = line
                    continue

                # Artikel
                m_art = artikel_re.search(line)
                if m_art:
                    current["Artikel"] = m_art.group(1)

                # PLZ / Ort
                if plz_re.search(line):
                    current["PLZ / Ort"] = line

                # Position Box + Adr.-Nr + Rhythmus
                m_pos = pos_re.search(line)
                if m_pos:
                    current["Position Box"] = m_pos.group(0)
                    parts = line.split()
                    if len(parts) >= 2:
                        current["Adr.-Nr."] = parts[-2]
                        current["Rhythmus"] = parts[-1]

                # Bemerkungen sammeln
                current["Bemerkung"] = (
                    current.get("Bemerkung", "") + " " + line
                ).strip()

    # letzten Datensatz sichern
    if current.get("Position Box"):
        rows.append(current)

    df = pd.DataFrame(rows)

    # Spalten sicherstellen
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

    for col in columns:
        if col not in df.columns:
            df[col] = ""

    df = df[columns]

    # nur echte Datensaetze behalten
    df = df[df["Position Box"].str.match(r"^\d+/\d+\.\d+$", na=False)]
    df.reset_index(drop=True, inplace=True)

    return df


# -----------------------------
# DataFrame -> Excel (Bytes)
# -----------------------------
def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Tour")
    return output.getvalue()


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(
    page_title="PDF → Excel (Tourenliste)",
    layout="centered",
)

st.title("PDF → Excel (Tourenliste)")
st.write(
    "PDF hochladen, automatisch in Excel umwandeln und direkt herunterladen. "
    "Keine Installation auf dem Arbeits-PC noetig."
)

uploaded = st.file_uploader("PDF Datei", type=["pdf"])

if uploaded:
    pdf_bytes = uploaded.read()

    with st.spinner("PDF wird verarbeitet …"):
        df = parse_tourenliste(pdf_bytes)

    st.success(f"Gefundene Eintraege: {len(df)}")

    if len(df) > 0:
        st.dataframe(df, use_container_width=True)

        excel_bytes = df_to_xlsx_bytes(df)
        filename = uploaded.name.replace(".pdf", ".xlsx")

        st.download_button(
            label="Excel herunterladen",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning("Keine Eintraege erkannt – Parser weiter anpassen.")
