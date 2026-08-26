# Word Doc Generator - FACTURA Generator

A simple local desktop app (Tkinter GUI) that fills in a Word (`.docx`)
template with details entered through a form, and saves a new generated document.
No internet connection or web server needed - runs entirely on your PC.

## Current status

This is the **FACTURA Generator** mode only. A second document type is planned
but disabled in the UI for now (radio button placeholder).

## How it works

1. You keep your existing Word template, but add unique placeholder tokens
   (like `{{NR}}`, `{{NUME}}`, `{{DATA}}`) at each spot where a value should
   be inserted. See `field_config.py` for the full list of tokens.
2. Put that edited template at `templates/factura_template.docx`.
3. Run the app, fill the form, click **GENERATE**.
4. A new `.docx` is created in the `output/` folder with your values
   inserted in place of the tokens, original template formatting preserved.
5. Fields defined as "uppercase+bold" are automatically converted to
   UPPERCASE and made bold. `profil_ales` and `tipul_geamului` are underlined.

## Setup

```bash
pip install -r requirements.txt
```

Place your prepared template at:

```
templates/factura_template.docx
```

Run:

```bash
python main.py
```

## Finding where to put placeholders in your template

If you're not sure where exactly the placeholder text needs to go inside
your existing `.docx`, run:

```bash
python inspect_template.py templates/factura_template.docx
```

This prints out every piece of text currently in the document (body, tables,
headers, footers), so you can locate the exact line/cell that needs a token,
open the file in Word, and type the token there (e.g. replace a blank
underscore line with `{{NR}}`).

## Fields (FACTURA Generator)

| UI Label | Field key | Placeholder token | Format |
|---|---|---|---|
| NR: | nr | `{{NR}}` | number |
| Data: | data | `{{DATA}}` | DD.MM.YYYY |
| NAME: | nume | `{{NUME}}` | text (UPPERCASE, bold) |
| Loc.: | loc | `{{LOC}}` | text (UPPERCASE, bold) |
| Adresa: | adresa | `{{ADRESA}}` | text (UPPERCASE, bold) |
| Tel: | tel | `{{TEL}}` | text (UPPERCASE, bold) |
| Cod fiscal: | cod_fiscal | `{{COD_FISCAL}}` | text (UPPERCASE, bold) |
| Inmatriculare la RC: | inmatriculare_rc | `{{INMATRICULARE_RC}}` | text (UPPERCASE, bold) |
| Cont: | cont | `{{CONT}}` | IBAN (UPPERCASE, bold) |
| Banca: | banca | `{{BANCA}}` | text (UPPERCASE, bold) |
| Reprezentata prin: | reprezentata_prin | `{{REPREZENTATA_PRIN}}` | text (UPPERCASE, bold) |
| CNP: | cnp | `{{CNP}}` | number (UPPERCASE, bold) |
| Anexa contract: | anexa_contract | `{{ANEXA_CONTRACT}}` | text (UPPERCASE, bold) |
| Termen de zile: | termen_zile | `{{TERMEN_ZILE}}` | number (UPPERCASE, bold) |
| Pana la data de: | pana_la_data | `{{PANA_LA_DATA}}` | DD.MM.YYYY (UPPERCASE, bold) |
| Profil ales: | profil_ales | `{{PROFIL_ALES}}` | text (underlined) |
| Tipul geamului: | tipul_geamului | `{{TIPUL_GEAMULUI}}` | text (underlined) |

## Building a standalone .exe (optional)

Once the app works, package it with PyInstaller:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name FacturaGenerator main.py
```

The `.exe` will be in `dist/FacturaGenerator.exe`. Make sure to also copy
the `templates/` folder next to the `.exe` (PyInstaller does not bundle it
automatically unless you add `--add-data`).

## Next step

Send over the template `.docx` and a screenshot of the layout - once shared,
the placeholder tokens can be inserted at the exact right spots and
`field_config.py` / `PLACEHOLDER_MAP` adjusted to match precisely.
