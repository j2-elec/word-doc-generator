"""
Field configuration for the FACTURA Generator (contract mode).

FACTURA_FIELDS: defines the UI form fields, in order, shown for this mode.
PLACEHOLDER_MAP: maps each field key to the literal placeholder text that
                  must exist inside templates/template_contract.docx.
UPPERCASE_BOLD_FIELDS: field keys whose value should be shown UPPERCASE + BOLD.
UNDERLINE_FIELDS: field keys whose value should be underlined.

These placeholder tokens match template_contract.docx exactly as provided
(bare words, no curly braces, already bold/underlined in the template itself).
"""

FACTURA_FIELDS = [
    {"key": "nr", "label": "NR:", "type": "number"},
    {"key": "data", "label": "Data:", "type": "date"},
    {"key": "nume", "label": "NAME:", "type": "text"},
    {"key": "loc", "label": "Loc.:", "type": "text"},
    {"key": "adresa", "label": "Adresa:", "type": "text"},
    {"key": "tel", "label": "Tel:", "type": "text"},
    {"key": "cod_fiscal", "label": "Cod fiscal:", "type": "text"},
    {"key": "inmatriculare_rc", "label": "Inmatriculare la RC:", "type": "text"},
    {"key": "cont", "label": "Cont (IBAN):", "type": "text"},
    {"key": "banca", "label": "Banca:", "type": "text"},
    {"key": "reprezentata_prin", "label": "Reprezentata prin:", "type": "text"},
    {"key": "cnp", "label": "CNP:", "type": "text"},
    {"key": "termen_livrare", "label": "Termen livrare (zile):", "type": "number", "section_break": True},
    {"key": "data_livrare", "label": "Data livrare (pana la):", "type": "date"},
    {"key": "curs_bnr", "label": "Curs BNR:", "type": "text"},
    {"key": "valoare_contract", "label": "Valoare contract:", "type": "text"},
    {"key": "avans_contract", "label": "Avans contract:", "type": "text"},
    {"key": "diferenta_contract", "label": "Diferenta contract:", "type": "text"},
    {"key": "profil_ales", "label": "Profil ales:", "type": "text", "section_break": True},
    {"key": "tipul_geamului", "label": "Tipul geamului:", "type": "text"},
]

PLACEHOLDER_MAP = {
    "nr": "NrContract", "data": "DataContract", "nume": "NumeCompanie",
    "loc": "LocalitateCompanie", "adresa": "AdresaCompanie", "tel": "TelCompanie",
    "cod_fiscal": "CodFiscalCompanie", "inmatriculare_rc": "NumarInmatriculare",
    "cont": "ContCompanie", "banca": "BancaCompanie",
    "reprezentata_prin": "ReprezentantCompanie", "cnp": "CNPReprezentantCompanie",
    "termen_livrare": "TermenLivrare", "data_livrare": "DataLivrare",
    "curs_bnr": "CursBNR", "valoare_contract": "ValoareContract",
    "avans_contract": "AvansContract", "diferenta_contract": "DiferentaContract",
    "profil_ales": "SerieProfilAles", "tipul_geamului": "TipGeamAles",
}

UPPERCASE_BOLD_FIELDS = {
    "nume", "loc", "adresa", "tel", "cod_fiscal", "inmatriculare_rc",
    "cont", "banca", "reprezentata_prin", "cnp", "termen_livrare",
    "data_livrare", "curs_bnr", "valoare_contract", "avans_contract",
    "diferenta_contract",
}

UNDERLINE_FIELDS = {"profil_ales", "tipul_geamului"}
