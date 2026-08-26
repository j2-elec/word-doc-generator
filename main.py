"""
FACTURA Generator - Word Document Filler
Local desktop application for generating documents from a Word template.

Requirements:
    pip install python-docx

Usage:
    python main.py
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import re
from datetime import datetime

try:
    from docx import Document
except ImportError:
    print("python-docx is not installed. Run: pip install python-docx")
    sys.exit(1)

from field_config import FACTURA_FIELDS, PLACEHOLDER_MAP, UPPERCASE_BOLD_FIELDS, UNDERLINE_FIELDS

TEMPLATE_PATH = "templates/factura_template.docx"
OUTPUT_DIR = "output"


def validate_date(value: str) -> bool:
    return bool(re.match(r"^\d{2}\.\d{2}\.\d{4}$", value))


def validate_number(value: str) -> bool:
    return value.strip().isdigit()


class FacturaGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Document Generator")
        self.geometry("560x780")
        self.resizable(False, False)

        self.mode_var = tk.StringVar(value="factura")
        self.entries = {}

        self._build_ui()

    def _build_ui(self):
        top_frame = ttk.Frame(self, padding=10)
        top_frame.pack(fill="x")

        ttk.Label(top_frame, text="Select document type:", font=("Segoe UI", 10, "bold")).pack(anchor="w")

        radio_frame = ttk.Frame(top_frame)
        radio_frame.pack(fill="x", pady=5)

        ttk.Radiobutton(
            radio_frame, text="FACTURA Generator", variable=self.mode_var,
            value="factura", command=self._render_fields
        ).pack(side="left", padx=(0, 20))

        ttk.Radiobutton(
            radio_frame, text="Second document type (coming soon)", variable=self.mode_var,
            value="other", state="disabled"
        ).pack(side="left")

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=10, pady=5)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)

        self.scroll_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=(10, 0))
        scrollbar.pack(side="right", fill="y")

        self.form_container = self.scroll_frame

        bottom_frame = ttk.Frame(self, padding=10)
        bottom_frame.pack(fill="x")

        self.generate_btn = tk.Button(
            bottom_frame, text="GENERATE", font=("Segoe UI", 14, "bold"),
            bg="#2e7d32", fg="white", activebackground="#1b5e20",
            height=2, command=self.on_generate
        )
        self.generate_btn.pack(fill="x")

        self.status_label = ttk.Label(bottom_frame, text="", foreground="#555")
        self.status_label.pack(pady=(6, 0))

        self._render_fields()

    def _render_fields(self):
        for widget in self.form_container.winfo_children():
            widget.destroy()
        self.entries.clear()

        mode = self.mode_var.get()
        if mode != "factura":
            return

        row = 0
        for field in FACTURA_FIELDS:
            key = field["key"]
            label_text = field["label"]
            field_type = field.get("type", "text")

            if field.get("section_break"):
                spacer = ttk.Label(self.form_container, text="")
                spacer.grid(row=row, column=0, pady=8)
                row += 1

            label = ttk.Label(self.form_container, text=label_text, width=22, anchor="w")
            label.grid(row=row, column=0, sticky="w", padx=10, pady=6)

            entry = ttk.Entry(self.form_container, width=38)
            entry.grid(row=row, column=1, sticky="w", padx=5, pady=6)

            if field_type == "date":
                hint = ttk.Label(self.form_container, text="DD.MM.YYYY", foreground="#888")
                hint.grid(row=row, column=2, sticky="w")

            self.entries[key] = entry
            row += 1

    def _collect_values(self):
        values = {}
        for field in FACTURA_FIELDS:
            key = field["key"]
            raw_value = self.entries[key].get().strip()

            if field.get("required", True) and not raw_value:
                raise ValueError(f"Field '{field['label']}' is required.")

            if field.get("type") == "date" and raw_value:
                if not validate_date(raw_value):
                    raise ValueError(f"Field '{field['label']}' must be in DD.MM.YYYY format.")

            if field.get("type") == "number" and raw_value:
                if not validate_number(raw_value):
                    raise ValueError(f"Field '{field['label']}' must contain only digits.")

            if key in UPPERCASE_BOLD_FIELDS:
                display_value = raw_value.upper()
            else:
                display_value = raw_value

            values[key] = display_value
        return values

    def on_generate(self):
        try:
            values = self._collect_values()
        except ValueError as e:
            messagebox.showerror("Invalid input", str(e))
            return

        if not os.path.exists(TEMPLATE_PATH):
            messagebox.showerror(
                "Template missing",
                f"Could not find template at:\n{TEMPLATE_PATH}\n\n"
                "Place your .docx template there before generating."
            )
            return

        try:
            output_path = self.generate_document(values)
        except Exception as e:
            messagebox.showerror("Generation failed", str(e))
            return

        self.status_label.config(text=f"Saved: {output_path}")
        messagebox.showinfo("Success", f"Document generated:\n{output_path}")

    def generate_document(self, values: dict) -> str:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        doc = Document(TEMPLATE_PATH)

        for key, value in values.items():
            placeholder = PLACEHOLDER_MAP.get(key)
            if not placeholder:
                continue
            bold = key in UPPERCASE_BOLD_FIELDS
            underline = key in UNDERLINE_FIELDS
            replace_placeholder_everywhere(doc, placeholder, value, bold=bold, underline=underline)

        nr_value = values.get("nr", "output")
        safe_nr = re.sub(r"[^A-Za-z0-9_\-]", "_", nr_value)
        filename = f"Factura_{safe_nr}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        output_path = os.path.join(OUTPUT_DIR, filename)
        doc.save(output_path)
        return output_path


def replace_placeholder_everywhere(doc, placeholder: str, value: str, bold: bool = False, underline: bool = False):
    """
    Replaces a {{PLACEHOLDER}} token wherever it appears in the document body,
    tables, and headers/footers, preserving surrounding run formatting.
    """
    targets = list(doc.paragraphs)
    for table in doc.tables:
        for row_ in table.rows:
            for cell in row_.cells:
                targets.extend(cell.paragraphs)

    for section in doc.sections:
        for header_footer in (section.header, section.footer):
            targets.extend(header_footer.paragraphs)
            for table in header_footer.tables:
                for row_ in table.rows:
                    for cell in row_.cells:
                        targets.extend(cell.paragraphs)

    for paragraph in targets:
        _replace_in_paragraph(paragraph, placeholder, value, bold, underline)


def _replace_in_paragraph(paragraph, placeholder: str, value: str, bold: bool, underline: bool):
    full_text = "".join(run.text for run in paragraph.runs)
    if placeholder not in full_text:
        return

    new_full_text = full_text.replace(placeholder, value)

    if paragraph.runs:
        template_run = paragraph.runs[0]
        for run in list(paragraph.runs):
            run.text = ""
        paragraph.runs[0].text = new_full_text
        if bold:
            paragraph.runs[0].bold = True
        if underline:
            paragraph.runs[0].underline = True
    else:
        run = paragraph.add_run(new_full_text)
        if bold:
            run.bold = True
        if underline:
            run.underline = True


if __name__ == "__main__":
    app = FacturaGeneratorApp()
    app.mainloop()
