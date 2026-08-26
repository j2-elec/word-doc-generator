# Word Doc Generator

A simple local desktop app (Tkinter GUI) that fills Word (`.docx`) templates
with data entered through a form, or read in bulk from an Excel source file.
No internet connection or web server needed - runs entirely on your PC.

## Modes (radio button)

1. **FACTURA Generator** - manual form entry, fills one template and saves
   one output document. See "FACTURA Generator" section below.
2. **SOMATII Generator** - batch mode, reads `source_excel.xlsx` and generates
   three sets of Word documents for every matching company. See "SOMATII
   Generator" section below.

## Setup

```bash
pip install -r requirements.txt
python main.py
```

---

## FACTURA Generator

Fills in a single Word template using manually entered form fields.

### How it works

1. Your Word template must contain placeholder tokens like `{{NR}}`,
   `{{NUME}}`, `{{DATA}}` at each spot where a value should be inserted
   (see `field_config.py` for the full list).
2. Place the prepared template at `templates/factura_template.docx`.
3. Run the app, select "FACTURA Generator", fill the form, click GENERATE.
4. A new `.docx` is created in `output/` with your values inserted, keeping
   the template's original formatting.
5. Fields marked "uppercase+bold" in `field_config.py` are automatically
   converted to UPPERCASE and bolded. `profil_ales` and `tipul_geamului`
   are underlined.

### Finding placeholder spots

```bash
python inspect_template.py templates/factura_template.docx
```

Prints every piece of existing text (body, tables, headers, footers) so you
can find where to type each `{{TOKEN}}` in Word.

---

## SOMATII Generator

Batch-generates three official documents from a single Excel data source:
individual **somatie** letters (one per company), a combined **envelope**
document (one printable page per company), and a combined **borderou**
(dispatch register) document with one table row per company.

### Required files (same folder as the app / .exe)

- `source_excel.xlsx` - source data workbook
- `template_somatie.docx` - somatie letter template
- `template_envelope-2.docx` - envelope template
- `template_borderou-3.docx` - borderou (dispatch register) template

### Source Excel format

Data is read from the **last sheet** in the workbook. Only rows where
column **L** equals `somam` (matches `somăm`, with or without diacritics,
case-insensitive) are included. Relevant columns:

| Column | Meaning |
|---|---|
| B | CompanyName |
| H | CompanyTotal (decimal, e.g. 1234.5 -> shown as `1.234,50`) |
| I | CompanyCUI |
| J | CompanyJ (Registrul Comertului number) |
| K | CompanyAddress (raw, see format below) |
| L | Filter column - only `somam`/`somăm` rows are processed |

**Raw address format in column K** (always comma-separated, judet and
localitate first, then street/number last, and may include extra
apartment/block info after the number):

```
JUD. MURES, TARGU MURES, STRADA O STRADA, NR. 28
JUD. CLUJ, CLUJ-NAPOCA, STR. AVIATORILOR, NR. 12, BL. A2, AP. 5
```

The app automatically reorders and reformats this for each document type
(see table below). Only the first letter of each word is capitalized
(Title Case), not full uppercase, except where noted.

### Output per document type

| Document | Output | Address format used |
|---|---|---|
| Somatie | `somatie_<CompanyName>.docx` - one file per company | `Strada, Nr., Localitate, Jud. Judet` e.g. `Str. O Strada, Nr. 28, Targu Mures, Jud. Mures` |
| Envelope | `envelope_toate_companiile.docx` - one combined file, one page per company | Line 1: `Localitate, Strada, Nr.` Line 2: `Jud. Judet`. `CompanyName` shown in full UPPERCASE |
| Borderou | `borderou_toate_companiile.docx` - one combined file, table filled row by row | `Localitate, Strada, Nr., Judet` |

Placeholders replaced in each template: `CompanyName`, `CompanyAddress`,
`CompanyCUI`, `CompanyJ`, `CompanyTotal` (somatie only). All formatting
(bold, etc.) already present around the placeholder text in the templates
is preserved.

### Running it

1. Put `source_excel.xlsx` and the three `template_*.docx` files in the
   same folder as `main.py` (or next to the `.exe`).
2. Run the app, select "SOMATII Generator".
3. Click GENERATE.
4. Watch the output log panel - it prints which company is being read and
   which document is being written, in real time, for debugging.
5. All 4+ output files (N somatie files + 1 envelope + 1 borderou) are
   saved in the same folder.

---

## Building a standalone .exe (optional)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name DocumentGenerator main.py
```

Copy the `templates/` folder (for FACTURA mode) and the `source_excel.xlsx`
+ `template_*.docx` files (for SOMATII mode) next to the generated `.exe` -
PyInstaller does not bundle data files automatically.

## Files

| File | Purpose |
|---|---|
| `main.py` | Tkinter GUI, both modes |
| `field_config.py` | FACTURA Generator field/placeholder definitions |
| `inspect_template.py` | Dumps template text to help locate placeholder spots |
| `address_utils.py` | Shared address parsing/reformatting logic for SOMATII |
| `somatie_generator.py` | SOMATII batch generation logic (Excel read + 3 doc writers) |
| `requirements.txt` | Python dependencies |
