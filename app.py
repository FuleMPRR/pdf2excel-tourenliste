import re
from io import BytesIO
from typing import List, Dict, Tuple, Optional

import pandas as pd
import pdfplumber
import streamlit as st


# --- Helpers: PDF -> rows using x-positions (works well for "text PDFs" like your Tourenliste) ---

HEADER_HINTS = [
    "Firma",
    "Ansprech-Person",
    "Telefon",
    "Strasse",
    "PLZ",
    "Artikel",
    "PositionBox",
    "Bemerkung",
    "Reihenf",
    "Adr.-Nr.",
    "Rhyt.",
]

POSITION_RE = re.compile(r"^\d+/\d+\.\d+$")  # e.g. 86/1.0


def _group_words_into_lines(words: List[Dict], y_tol: float = 2.5) -> List[List[Dict]]:
    """Group extracted words into visual lines by 'top' coordinate."""
    if not words:
        return []

    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines: List[List[Dict]] = []
    current: List[Dict] = [words[0]]
    current_top = words[0]["top"]

    for w in words[1:]:
        if abs(w["top"] - current_top) <= y_tol:
            current.append(w)
        else:
            lines.append(sorted(current, key=lambda ww: ww["x0"]))
            current = [w]
            current_top = w["top"]

    lines.append(sorted(current, key=lambda ww: ww["x0"]))
    return lines


def _find_header_line(lines: List[List[Dict]]) -> Optional[int]:
    """Find line index that looks like the table header."""
    for i, line in enumerate(lines[:30]):  # header should be near top
        text = " ".join(w["text"] for w in line)
        hits = sum(1 for h in HEADER_HINTS if h in text)
        if hits >= 4 and "Firma" in text and ("Adr.-Nr." in text or "Rhyt." in text):
            return i
    return None


def _build_column_boundaries(header_line: List[Dict]) -> List[Tuple[str, float]]:
    """
    Build rough column anchors based on header word x positions.
    Returns list of (col_name, x0_anchor) sorted by x0_anchor.
    """
    # Normalize header tokens (some headers may be split)
    tokens = [(w["text"], w["x0"]) for w in header_line]

    def anchor_for(label: str) -> Optional[float]:
        # Find first token that contains label (or part of it)
        for t, x in tokens:
            if label in t:
                return x
        return None

    anchors: List[Tuple[str, float]] = []

    # Minimal columns you want in Excel (fits your shown result)
    col_map = {
        "Firma": anchor_for("Firma"),
        "Ansprechperson": anchor_for("Ansprech"),
        "Telefon": anchor_for("Telefon"),
        "Strasse": anchor_for("Strasse"),
        "PLZ / Ort": anchor_for("PLZ"),
        "Artikel": anchor_for("Artikel"),
        "Bemerkung": anchor_for("Bemerkung"),
        "Position Box": anchor_for("PositionBox"),
        "Adr.-Nr.": anchor_for("Adr.-Nr."),
        "Rhythmus": anchor_for("Rhyt."),
    }

    # Some PDFs might not match perfectly; keep only found anchors
    for name, x in col_map.items():
        if x is not None:
            anchors.append((name, x))

    anchors.sort(key=lambda a: a[1])
    return anchors


def _assign_words_to_columns(line: List[Dict], anchors: List[Tuple[str, float]]) -> Dict[str, str]:
    """
    Assign each word to nearest column region based on x0 anchors.
    Simple heuristic: each word goes to the last anchor whose x0 <= word.x0.
    """
    out = {name: "" for name, _ in anchors}
    if not anchors:
        return out

    for w in line:
        wx = w["x0"]
        # find last anchor with x0 <= wx
        idx = 0
        for j, (_, ax) in enumerate(anchors):
            if wx >= ax:
                idx = j
            else:
                break
        col = anchors[idx][0]
        out[col] = (out[col] + " " + w["text"]).strip()

    return out


def parse_tourenliste(pdf_bytes: bytes) -> pd.DataFrame:
    rows: List[Dict[str, str]] = []

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        anchors: Optional[List[Tuple[str, float]]] = None

        for page in pdf.pages:
            words = page.extract_words(
                keep_blank_chars=False,
                use_text_flow=True,
            )
            lines = _group_words_into_lines(words)

            header_idx = _find_header_line(lines)
            if header_idx is not None and anchors is None:
                anchors = _build_column_boundaries(lines[header_idx])

            # If we still didn't detect anchors, skip (or fallback to simple text parsing)
            if not anchors:
                continue

            # Process lines below header
            buffer: Dict[str, str] = {name: "" for name, _ in anchors}

            for line in lines[(header_idx + 1) if header_idx is not None else 0 :]:
                line_text = " ".join(w["text"] for w in line).strip()

                # Skip empty or footer lines
                if not line_text:
                    continue
                if "Seite:" in line_text or line_text.startswith("103_Tourenliste"):
                    continue
                if line_text.startswith("Tour ") or line_text.startswith("Tourenliste"):
                    continue

                assigned = _assign_words_to_columns(line, anchors)

                # Merge into buffer (handles multi-line remarks, addresses, etc.)
                for k, v in assigned.items():
                    if v:
                        if buffer.get(k):
                            buffer[k] = (buffer[k] + " " + v).strip()
                        else:
                            buffer[k] = v.strip()

                # Decide when a logical record ends: Position Box present (e.g. 86/12.0)
                pos_val = buffer.get("Position Box", "").strip()
                if pos_val and POSITION_RE.match(pos_val.split()[-1]):
                    # Cleanup: keep only last token for Position Box if extra text got in
                    buffer["Position Box"] = pos_val.split()[-1]

                    # Try to tidy phone fields: remove stray slashes spacing
                    if "Telefon" in buffer:
                        buffer["Telefon"] = buffer["Telefon"].replace(" / ", " / ").strip()

                    rows.append(buffer.copy())
                    buffer = {name: "" for name, _ in anchors}

    # Ensure a stable column order
    wanted_cols = [
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
    # Only keep columns that exist (in case anchors missed something)
    existing_cols = [c for c in wanted_cols if any(c == a[0] for a in (anchors or []))]
    df = pd.DataFrame(rows)

    # If some columns were not detected, still include them as empty
    for c in wanted_cols:
        if c not in df.columns:
            df[c] = ""

    df = df[wanted_cols]

    # Final cleanup: remove double spaces
    df = df.applymap(lambda x: re.sub(r"\s{2,}", " ", x).strip() if isinstance(x, str) else x)

    # Remove obviously wrong rows (e.g. totals line)
    if "Position Box" in df.columns:
        df = df[df["Position Box"].str.match(POSITION_RE, na=False)]

    df.reset_index(drop=True, inplace=True)
    return df


def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Tour")
    return bio.getvalue()


# --- Streamlit UI ---

st.set_page_config(page_title="PDF → Excel (Tourenliste)", layout="centered")
st.title("PDF → Excel (Tourenliste)")

st.write("PDF hochladen, dann Excel direkt herunterladen. Keine Installation auf dem Arbeits-PC noetig.")

uploaded = st.file_uploader("PDF Datei", type=["pdf"])

if uploaded is not None:
    pdf_bytes = uploaded.read()

    with st.spinner("Konvertiere…"):
        df = parse_tourenliste(pdf_bytes)

    st.success(f"Gefundene Eintraege: {len(df)}")
    st.dataframe(df, use_container_width=True)

    xlsx_bytes = df_to_xlsx_bytes(df)
    out_name = uploaded.name.rsplit(".", 1)[0] + ".xlsx"

    st.download_button(
        label="Excel herunterladen",
        data=xlsx_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
