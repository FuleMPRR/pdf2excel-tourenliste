def parse_tourenliste(pdf_bytes: bytes) -> pd.DataFrame:
    import re
    from io import BytesIO

    rows = []
    current = {}

    pos_re = re.compile(r"\b\d+/\d+\.\d+\b")           # z.B. 86/12.0
    plz_re = re.compile(r"\b\d{4}\b")                  # CH PLZ
    artikel_re = re.compile(r"\b(KB|GB 2023|DGB 2023)\b")

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = [l.strip() for l in text.split("\n") if l.strip()]

            for line in lines:
                # Footer/Headers ignorieren
                if "Tourenliste per:" in line or line.startswith("103_Tourenliste") or "Seite:" in line:
                    continue
                if line.startswith("Tour "):
                    continue

                # Start eines neuen Datensatzes: Firmenzeile (heuristisch: keine Positionsbox drin)
                if (not pos_re.search(line)
                        and not line[:1].isdigit()
                        and "Woche" not in line
                        and "Firma" not in line
                        and len(line) > 2):
                    # vorherigen Datensatz abschliessen, falls vorhanden
                    if current.get("Position Box"):
                        rows.append(current)
                        current = {}

                    current["Firma"] = line
                    continue

                # Artikel
                m_art = artikel_re.search(line)
                if m_art:
                    current["Artikel"] = m_art.group(1)

                # PLZ/Ort
                if plz_re.search(line) and ("PLZ / Ort" not in current):
                    current["PLZ / Ort"] = line

                # Positionsbox + Adr.-Nr + Rhythmus (steht am Ende der Zeile)
                m_pos = pos_re.search(line)
                if m_pos:
                    current["Position Box"] = m_pos.group(0)
                    parts = line.split()
                    if len(parts) >= 2:
                        current["Adr.-Nr."] = parts[-2]
                        current["Rhythmus"] = parts[-1]

                # Bemerkung sammeln (alles, was nicht sauber in andere Felder faellt)
                current["Bemerkung"] = (current.get("Bemerkung", "") + " " + line).strip()

    # letzten Datensatz abschliessen
    if current.get("Position Box"):
        rows.append(current)

    df = pd.DataFrame(rows)
    for col in ["Firma", "Ansprechperson", "Telefon", "Strasse", "PLZ / Ort", "Artikel",
                "Bemerkung", "Position Box", "Adr.-Nr.", "Rhythmus"]:
        if col not in df.columns:
            df[col] = ""

    df = df[["Firma", "Ansprechperson", "Telefon", "Strasse", "PLZ / Ort", "Artikel",
             "Bemerkung", "Position Box", "Adr.-Nr.", "Rhythmus"]]

    # Nur echte Datensaetze behalten
    df = df[df["Position Box"].str.match(r"^\d+/\d+\.\d+$", na=False)].reset_index(drop=True)
    return df
