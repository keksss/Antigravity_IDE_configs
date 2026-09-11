---
name: markdown-to-docx
description: >-
  Generates professional Microsoft Word (.docx) documents from Markdown (.md) files.
  Supports full typography, tables, lists, code blocks, embedded images, and custom styling
  via reference .docx templates. Use when the user asks to "create docx from markdown",
  "convert md to docx", "generate word document from md", or "apply docx template".
---

# markdown-to-docx Skill

Convert Markdown (`.md`) files into polished Microsoft Word (`.docx`) documents with support for embedded images, tables, hyperlinks, code formatting, and custom corporate templates.

---

## Environment Setup

This skill uses Python with `python-docx` and `markdown-it-py`. Dependencies are declared via PEP 723 metadata in the script header.

Execute scripts using either:
1. **With `uv` (Recommended):**
   ```bash
   uv run scripts/md_to_docx.py <path_to_md>
   ```
2. **With Python and local virtual environment:**
   If `.venv` does not exist in this skill folder, create it:
   ```bash
   # Windows
   python -m venv .venv
   .\.venv\Scripts\pip install python-docx markdown-it-py
   .\.venv\Scripts\python scripts/md_to_docx.py <path_to_md>

   # Linux/macOS
   python -m venv .venv
   ./.venv/bin/pip install python-docx markdown-it-py
   ./.venv/bin/python scripts/md_to_docx.py <path_to_md>
   ```

---

## Usage Scenarios & Commands

### 1. Basic Conversion (Clean Default Styling)
```bash
uv run scripts/md_to_docx.py document.md
```
- Produces: `document.docx`
- Applies clean, modern typography (Calibri, 11pt, standard margins).

### 2. Using an Existing DOCX as a Styling Template
When the user wants the new document to match an existing Word document or corporate template:
```bash
uv run scripts/md_to_docx.py document.md --template templates/corporate_style.docx -o output/final_report.docx
```
- **How Template Inheritance Works:**
  - The script opens `corporate_style.docx`.
  - Retains all styles (`Heading 1`, `Normal`, font families, color themes, headers, footers, margins).
  - Clears existing body content and renders the Markdown using the template's styles.

### 3. Including Images and Tables
If `document.md` references relative images like `![Chart](images/q1_chart.png)`, the script automatically locates and embeds them centered within the generated Word document.

---

## CLI Options Reference

| Argument | Description | Default |
| :--- | :--- | :--- |
| `md_file` | Path to the source Markdown document. | *(Required)* |
| `-o`, `--output` | Destination `.docx` file path. | `<md_name>.docx` |
| `-t`, `--template` | Path to a reference `.docx` file for style inheritance. | `None` (Default styles) |
| `--engine` | Conversion engine: `auto`, `python` (enhanced AST renderer), or `pandoc`. | `auto` |

---

## Supported Markdown Elements (100% Syntax Coverage)

- **Headings:** `#` to `######` mapped directly to Word styles `Heading 1` through `Heading 6` (including Russian/Cyrillic headings).
- **Inline Formatting:** `**bold**`, `*italic*`, `***bold italic***`, `~~strikethrough~~`, `` `inline code` ``, and arbitrary nested combinations.
- **Line Breaks:** Proper `softbreak` (inter-word spacing preservation) and `hardbreak` (`run.add_break()` for two spaces / backslash line breaks).
- **Hyperlinks:** Native Word clickable hyperlinks (`w:hyperlink`) supporting formatted anchor text (`[**link**](url)`).
- **Lists:** Bulleted (`-`, `*`) and ordered (`1.`, `2.`) lists with multi-level nested sub-items (3+ tiers).
- **Task Lists (Checkboxes):** `- [ ]` and `- [x]` / `- [X]` rendered with clean check symbols (`☐` and `☑`).
- **Callouts / Alerts:** GitHub-style `> [!NOTE]`, `> [!TIP]`, `> [!IMPORTANT]`, `> [!WARNING]`, `> [!CAUTION]` rendered with colored accent borders, background shading, and full inline markdown formatting inside.
- **Blockquotes:** Standard and multi-paragraph blockquotes rendered with `Quote` styling and indentation.
- **Code Blocks:** Multi-line fenced blocks rendered with light gray background shading, left accent border, and Consolas font.
- **Tables:** GFM tables with column alignments (left, center, right), repeated header rows (`tblHeader`), `cantSplit` protection against page breaks, and rich inline formatting in cells.
- **Images:** Embedded and centered proportionally, preserving natural aspect ratio and DPI using Pillow.
- **Horizontal Rules:** Centered divider line (`---`, `***`).

---

## Chained Workflow: Post-Processing & Fine-Tuning with `docx-editor`

`markdown-to-docx` is strictly responsible for structural Markdown-to-DOCX conversion. It does not attempt to be a full OpenXML desktop publishing engine.

When you need:
- Custom corporate layout adjustments or color scheme modifications
- In-place text replacements or header/footer fine-tuning
- Inserting additional sections, table rows, or dynamic template elements

**Always chain the workflow:**
1. Generate the initial Word document from Markdown:
   ```bash
   uv run scripts/md_to_docx.py document.md -o output/report.docx
   ```
2. Switch to the `docx-editor` skill to inspect and perform fine-tuning:
   ```bash
   # Inspect available styles
   uv run ../docx-editor/scripts/docx_inspect.py output/report.docx --mode styles

   # Fine-tune content or replace text in-place
   uv run ../docx-editor/scripts/docx_edit.py output/report.docx replace-text --find "DRAFT" --replace "APPROVED"
   ```
