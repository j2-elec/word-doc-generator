"""
Address parsing/formatting utilities shared by the SOMATII batch generator.

Source format (column K in source_excel.xlsx), always comma-separated, e.g.:
    JUD. MURES, TARGU MURES, STRADA O STRADA, NR. 28
    JUD. MURES, TARGU MURES, STRADA O STRADA, NR. 28, BL. A2, AP. 5

Parsing rule:
    part[0]  -> judet
    part[1]  -> localitate
    part[2:] -> strada + numarul (+ any extra bloc/apartament info), kept
                joined together in original order since the "numarul" is
                always the last meaningful token of this group.
"""

import re


def title_case_ro(text: str) -> str:
    """Title-case a string but keep short abbreviations (Str., Nr., Bl.) sane."""
    text = text.strip().lower()
    words = text.split(" ")
    out = []
    for w in words:
        if not w:
            continue
        if "." in w:
            out.append(".".join(part.capitalize() for part in w.split(".")))
        elif "-" in w:
            out.append("-".join(part.capitalize() for part in w.split("-")))
        else:
            out.append(w.capitalize())
    return " ".join(out)


def parse_address(raw: str):
    """
    Splits the raw source address into (judet, localitate, strada_numarul).
    strada_numarul preserves everything after judet+localitate, comma-joined.
    """
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    if len(parts) < 3:
        judet = parts[0] if len(parts) > 0 else ""
        localitate = parts[1] if len(parts) > 1 else ""
        strada_numarul = ""
    else:
        judet = parts[0]
        localitate = parts[1]
        strada_numarul = ", ".join(parts[2:])

    judet_clean = re.sub(r"^jud\.?\s*", "", judet, flags=re.IGNORECASE).strip()

    return {
        "judet_raw": judet,
        "judet": title_case_ro(judet_clean),
        "localitate": title_case_ro(localitate),
        "strada_numarul": title_case_ro(strada_numarul),
    }


def format_address_somatie(raw: str) -> str:
    """
    Str. O Strada, Nr. 28, Targu Mures, Jud. Mures
    """
    a = parse_address(raw)
    return f"{a['strada_numarul']}, {a['localitate']}, Jud. {a['judet']}"


def format_address_envelope(raw: str):
    """
    Returns (line1, line2) for the envelope:
    line1: Localitate, Strada, Numarul
    line2: Jud. Judet
    """
    a = parse_address(raw)
    line1 = f"{a['localitate']}, {a['strada_numarul']}"
    line2 = f"Jud. {a['judet']}"
    return line1, line2


def format_address_borderou(raw: str) -> str:
    """
    Localitatea, Strada, Numarul, Judet
    """
    a = parse_address(raw)
    return f"{a['localitate']}, {a['strada_numarul']}, {a['judet']}"
