---
name: docx-to-pdf
description: >-
  Converts Microsoft Word (.docx) documents into professional PDF files with high-fidelity
  layout, font, and style preservation. Supports single files, custom destination paths,
  and batch directory conversion using native Microsoft Word COM automation or LibreOffice headless.
  Use when the user asks to "convert docx to pdf", "export word to pdf", "save docx as pdf",
  "generate pdf from docx", or "batch convert docx to pdf".
---

# docx-to-pdf Skill

Convert `.docx` documents into high-quality PDF files with complete preservation of pagination, embedded images, tables, vector shapes, headers, footers, and typography.

---

## Environment Setup

This skill uses Python with `docx2pdf` (for native Microsoft Word COM automation on Windows/macOS) and supports headless LibreOffice CLI as an alternative or cross-platform engine. Dependencies are declared via PEP 723 metadata in the script header.

Execute scripts using either:
1. **With `uv` (Recommended):**
   ```bash
   uv run scripts/docx_to_pdf.py <path_to_docx>
   ```
2. **With Python and local virtual environment:**
   If `.venv` does not exist in this skill folder, create it:
   ```bash
   # Windows
   python -m venv .venv
   .\.venv\Scripts\pip install python-docx docx2pdf
   .\.venv\Scripts\python scripts/docx_to_pdf.py <path_to_docx>

   # Linux/macOS
   python -m venv .venv
   ./.venv/bin/pip install python-docx docx2pdf
   ./.venv/bin/python scripts/docx_to_pdf.py <path_to_docx>
   ```

---

## Conversion Backends

The script automatically detects and selects the best available engine:

1. **Microsoft Word (COM Automation)** *(Default on Windows/macOS if Word is installed)*:
   - Provides 100% native Word rendering fidelity, identical to printing from Word.
2. **LibreOffice Headless (`soffice`)** *(Cross-platform for Windows, Linux, macOS)*:
   - Used when Microsoft Word is not installed, or when executing in CI/CD and Linux environments.
   - Automatically detected in system `PATH` and standard install directories (`C:\Program Files\LibreOffice\...`, `/usr/bin/soffice`, `/Applications/LibreOffice.app`).
   - Can be explicitly pointed to via the `LIBREOFFICE_PATH` environment variable.

---

## Usage Scenarios & Commands

### 1. Default Conversion (Generates `.pdf` next to `.docx`)
```bash
uv run scripts/docx_to_pdf.py document.docx
# Produces: document.pdf
```

### 2. Custom Output File
```bash
uv run scripts/docx_to_pdf.py reports/annual_2026.docx -o exports/annual_2026.pdf
```

### 3. Batch Directory Conversion
Convert all `.docx` files in a folder and place generated PDFs into a target directory:
```bash
uv run scripts/docx_to_pdf.py documents/ -o pdf_archive/
```

### 4. Specifying the Conversion Engine
Force the use of a specific rendering backend:
```bash
# Force Microsoft Word
uv run scripts/docx_to_pdf.py document.docx --engine word

# Force LibreOffice
uv run scripts/docx_to_pdf.py document.docx --engine libreoffice
```

### 5. Overwrite Existing PDF
```bash
uv run scripts/docx_to_pdf.py document.docx --force
```

---

## CLI Options Reference

| Argument | Description | Default |
| :--- | :--- | :--- |
| `inputs` | One or more `.docx` files or directories containing `.docx` documents. | *(Required)* |
| `-o`, `--output` | Destination `.pdf` file path (for single file) or destination directory (for batch mode). | `<input_name>.pdf` |
| `-e`, `--engine` | Backend engine: `auto`, `word`, or `libreoffice`. | `auto` |
| `-f`, `--force` | Overwrite existing output PDF file(s) without prompting. | `False` |

---

## Troubleshooting

- **No conversion engine succeeded:** Ensure either Microsoft Word is installed (Windows/macOS) or LibreOffice is installed.
- **File lock errors:** If the document is open in Microsoft Word, close the file in Word before converting to avoid access conflicts. Temporary Word lock files (`~$*.docx`) are automatically ignored.
