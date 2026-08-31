# Word Doc Generator

A simple local desktop app (Tkinter GUI) that fills Word (`.docx`) templates with data entered through a form, or read in bulk from an Excel source file. No internet connection or web server is needed.

## Modes

1. **FACTURA Generator** — manually fills `templates/template_contract.docx` and saves one contract document.
2. **SOMATII Generator** — reads `source_excel.xlsx` and creates individual somații, a combined envelope document, and a combined borderou.

## Setup

```bash
pip install -r requirements.txt
python main.py
```

## FACTURA Generator

Place the current contract template at:

```text
templates/template_contract.docx
```

The template must contain these literal placeholder tokens (no curly braces):

| UI label | Token in contract template |
|---|---|
| NR | `NrContract` |
| Data | `DataContract` |
| Nume | `NumeCompanie` |
| Loc. | `TownCompany` |
| Adresa | `AdressCompany` |
| Tel | `PhoneCompany` |
| Cod fiscal | `FiscalCodeCompany` |
| Inmatriculare la RC | `NumberRCCompany` |
| Cont | `ContCompanie` |
| Banca | `BankCompany` |
| Reprezentata prin | `RepresentativeCompany` |
| CNP | `CNPReprezentantCompanie` |
| Termen livrare | `TermenLivrare` |
| Data livrare | `DataLivrare` |
| Curs BNR | `CursBNR` |
| Valoare contract | `ValoareContract` |
| Avans contract | `AvansContract` |
| Diferenta contract | `DiferentaContract` |
| Profil ales | `SerieProfilAles` |
| Tipul geamului | `TipGeamAles` |

The program replaces longer tokens first. This is necessary because `ReprezentantCompanie` is contained inside `CNPReprezentantCompanie`; replacing in the opposite order would corrupt the CNP placeholder.

Generated files are saved under `output/` as `NumeCompanie_CT-Vanzare-Cumparare-NICOMY-NDY.docx`.

## SOMATII Generator

Keep these files in the same folder as `main.py` (or beside the packaged `.exe`):

- `source_excel.xlsx`
- `template_somatie.docx`
- `template_envelope-2.docx`
- `template_borderou-3.docx`

The program reads the last Excel sheet and includes only rows where column L is `somam` or `somăm` (case-insensitive). It uses B=CompanyName, H=CompanyTotal, I=CompanyCUI, J=CompanyJ, K=CompanyAddress and L=filter.

Outputs are individual `somatie_<CompanyName>.docx` files plus `envelope_toate_companiile.docx` and `borderou_toate_companiile.docx`.

## Build an EXE

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name DocumentGenerator main.py
```

Copy the required templates and Excel source next to the `.exe`; PyInstaller does not include data files by default.
