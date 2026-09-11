# docx-tools Plugin for Google Antigravity

A modular, portable Antigravity plugin for working with Microsoft Word (`.docx`) documents. Provides full text/media extraction, template-driven compilation from Markdown, and safe in-place editing with strict style preservation.

---

## Skills in this Plugin

This plugin follows the **Single Responsibility** principle and is split into three focused skills:

| Skill | Directory | Description |
| :--- | :--- | :--- |
| **`docx-to-markdown`** | [`skills/docx-to-markdown/`](./skills/docx-to-markdown/SKILL.md) | Converts `.docx` to `.md`, extracting all images to a dedicated folder and preserving hyperlinks and tables. Supports Microsoft `markitdown` and `mammoth` engines. |
| **`markdown-to-docx`** | [`skills/markdown-to-docx/`](./skills/markdown-to-docx/SKILL.md) | Compiles `.md` into a styled `.docx`, with rich formatting in table cells, code block syntax shading, GitHub alerts (`[!NOTE]`), nested lists, and reference templates (`--template`). |
| **`docx-editor`** | [`skills/docx-editor/`](./skills/docx-editor/SKILL.md) | Inspects document structure, fills Jinja2 templates (`docxtpl`), and performs safe in-place edits (text replacement, appending, table updates) while preserving existing styles, themes, and layouts. |
| **`docx-to-pdf`** | [`skills/docx-to-pdf/`](./skills/docx-to-pdf/SKILL.md) | Converts `.docx` documents to high-fidelity PDF using native Word COM (with hang protection) or LibreOffice headless, supporting single and batch conversion. |

---

## Integrity and Safety Rules

The plugin includes a mandatory safety rule:
- [`rules/docx-integrity.md`](./rules/docx-integrity.md): Enforces automatic backups (`.bak`), style inheritance, and OpenXML integrity protection.

---

## Directory Structure

```text
docx-tools/
├── plugin.json                       # Plugin manifest
├── README.md                         # This documentation
├── rules/
│   └── docx-integrity.md             # Style preservation & backup rules
└── skills/
    ├── docx-to-markdown/
    │   ├── SKILL.md
    │   └── scripts/
    │       └── docx_to_md.py         # Multi-engine converter (MarkItDown / Mammoth)
    ├── markdown-to-docx/
    │   ├── SKILL.md
    │   └── scripts/
    │       └── md_to_docx.py         # Rich markdown compiler with Pillow DPI scaling
    ├── docx-editor/
    │   ├── SKILL.md
    │   └── scripts/
    │       ├── docx_inspect.py       # Safe XML structure, styles & grid inspection
    │       ├── docx_edit.py          # Atomic run-aware text replacement & appending
    │       └── docx_render.py        # Jinja2 template rendering via docxtpl
    └── docx-to-pdf/
        ├── SKILL.md
        └── scripts/
            └── docx_to_pdf.py        # Subprocess-isolated Word COM & LibreOffice PDF export
```

---

## Quick Usage Guide

### 1. Unpack DOCX to Markdown and Images
```bash
# High-fidelity conversion via Microsoft MarkItDown
uv run skills/docx-to-markdown/scripts/docx_to_md.py document.docx --engine markitdown

# Legacy/semantic conversion via Mammoth
uv run skills/docx-to-markdown/scripts/docx_to_md.py document.docx --engine mammoth
```

### 2. Generate DOCX from Markdown (with rich elements)
```bash
# Clean default styling (tables with bold/code, callout alerts, code blocks, nested lists)
uv run skills/markdown-to-docx/scripts/md_to_docx.py notes.md

# Styled with an existing corporate template
uv run skills/markdown-to-docx/scripts/md_to_docx.py report.md --template templates/corporate.docx -o output/report.docx
```

### 3. Inspect and Safely Edit Existing DOCX
```bash
# Inspect headings, styles, and tables safely
uv run skills/docx-editor/scripts/docx_inspect.py spec.docx --mode outline

# Replace text safely across all runs, headers, and tables
uv run skills/docx-editor/scripts/docx_edit.py spec.docx replace-text --find "Версия 1.0" --replace "Версия 2.0"

# Append text matching the document's style
uv run skills/docx-editor/scripts/docx_edit.py spec.docx append-paragraph --text "Финальное примечание."
```

### 4. Template-Based Document Generation (docxtpl)
```bash
# Inspect undeclared variables in a Jinja2 template
uv run skills/docx-editor/scripts/docx_render.py template.docx --inspect-vars

# Render template using JSON data file (supports table loops, ifs, images)
uv run skills/docx-editor/scripts/docx_render.py template.docx -d context.json -o output/invoice.docx
```

### 5. Convert DOCX to PDF
```bash
# Default conversion (Word COM with 20s timeout or LibreOffice headless)
uv run skills/docx-to-pdf/scripts/docx_to_pdf.py report.docx

# Custom output path or batch folder conversion
uv run skills/docx-to-pdf/scripts/docx_to_pdf.py reports/ -o pdfs/
```

---

## Portability

This plugin uses strictly relative paths. To make it globally available across all your Antigravity workspaces, simply copy the `docx-tools` folder:
- **From:** `.agents/plugins/docx-tools/`
- **To:** `~/.gemini/config/plugins/docx-tools/`
