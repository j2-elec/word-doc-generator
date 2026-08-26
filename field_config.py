"""
Field configuration for the FACTURA Generator.

FACTURA_FIELDS: defines the UI form fields, in order, for the FACTURA generator.
PLACEHOLDER_MAP: maps each field key to the placeholder token that must exist
                  inside the Word template (templates/factura_template.docx).
UPPERCASE_BOLD_FIELDS: field keys whose value should be shown UPPERCASE + BOLD.
UNDERLINE_FIELDS: field keys whose value should be underlined.

IMPORTANT:
Open your .docx template in Word and make sure each corresponding spot
contains the exact placeholder text shown below (including double curly braces),
e.g. type {{NR}} where the invoice number should appear.
You can format the placeholder text itself however you like (font, size) -
the script will insert plain text into that same run's paragraph and then
apply bold/underline only for the fields listed below.
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

    {"key": "anexa_contract", "label": "Anexa contract:", "type": "text", "section_break": True},
    {"key": "termen_zile", "label": "Termen de zile:", "type": "number"},
    {"key": "pana_la_data", "label": "Pana la data de:", "type": "date"},

    {"key": "profil_ales", "label": "Profil ales:", "type": "text", "section_break": True},
    {"key": "tipul_geamului", "label": "Tipul geamului:", "type": "text"},
]

PLACEHOLDER_MAP = {
    "nr": "{{NR}}",
    "data": "{{DATA}}",
    "nume": "{{NUME}}",
    "loc": "{{LOC}}",
    "adresa": "{{ADRESA}}",
    "tel": "{{TEL}}",
    "cod_fiscal": "{{COD_FISCAL}}",
    "inmatriculare_rc": "{{INMATRICULARE_RC}}",
    "cont": "{{CONT}}",
    "banca": "{{BANCA}}",
    "reprezentata_prin": "{{REPREZENTATA_PRIN}}",
    "cnp": "{{CNP}}",
    "anexa_contract": "{{ANEXA_CONTRACT}}",
    "termen_zile": "{{TERMEN_ZILE}}",
    "pana_la_data": "{{PANA_LA_DATA}}",
    "profil_ales": "{{PROFIL_ALES}}",
    "tipul_geamului": "{{TIPUL_GEAMULUI}}",
}

UPPERCASE_BOLD_FIELDS = {
    "nume", "loc", "adresa", "tel", "cod_fiscal", "inmatriculare_rc",
    "cont", "banca", "reprezentata_prin", "cnp",
    "anexa_contract", "termen_zile", "pana_la_data",
}

UNDERLINE_FIELDS = {
    "profil_ales", "tipul_geamului",
}
