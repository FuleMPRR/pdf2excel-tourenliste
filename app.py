def parse_tourenliste(pdf_bytes: bytes) -> pd.DataFrame:
    import re
    rows = []
    current = {}

    pos_re = re.compile(r"\d+/\d+\.\d+")
    plz_re = re.compile(r"\b\d{4}\b")
    artikel_re = re.compile(r"\b(KB|GB 2023|DGB 2023)\b")

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            lines = [l.strip() for l in text.split("\n") if l.strip()]

            for line in lines:
                # neuer Datensatz beginnt meist mit Firmenname (keine Nummer am Anfang)
                if not re.match(r"^\d", line) and "/" not in line and "Tour" not in line:
                    if current and "Position Box" in current:
                        rows.append(current)
                        current = {}

                    current["Firma"] = line
                    continue

                if "Telefon" not in current and "+" in line:
                    current["Telefon"] = line
                if "Strasse" not in current and "/" in line:
                    current["Strasse"] = line.split("/")[0].strip()

                if plz_re.search(line):
                    current["PLZ / Ort"] = line

                if artikel_re.search(line):
