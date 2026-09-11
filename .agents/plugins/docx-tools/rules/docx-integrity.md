---
trigger: always_on
---

# DOCX Document Integrity and Safety Rules

When reading, modifying, or creating Microsoft Word (`.docx`) files using the `docx-tools` plugin, adhere to the following safety and style preservation guidelines.

---

## 1. Safety and Automatic Backup

1. **Always Backup Before Editing:**
   - Never perform destructive in-place writes to a `.docx` file without creating a `.bak` backup first.
   - The editing scripts (`docx_edit.py`) create backups automatically unless explicitly disabled with `--no-backup`. Keep backups enabled for safety.

2. **Verify File Existence and Integrity:**
   - Before editing, run the inspection script or check that the file is a valid `.docx` archive (ZIP format containing `word/document.xml`).
   - If a file is open in Microsoft Word, Word may place an exclusive lock on the file. If an access error occurs, advise the user or handle file locks gracefully.

---

## 2. Style and Layout Preservation

1. **Preserve OpenXML Elements:**
   - A `.docx` document is not raw text; it contains complex style definitions (`styles.xml`), document themes, section properties (margins, headers, footers), and table formatting.
   - Never convert a `.docx` file to plain text or Markdown and then attempt to re-save it as the original `.docx`. This completely destroys custom corporate formatting, styles, headers, and metadata.
   - For in-place editing, always use the `docx-editor` skill which operates directly on the OpenXML DOM.

2. **Style Inheritance When Appending Content:**
   - When adding new paragraphs or headings, match the style names already present in the document (e.g., `Heading 1`, `Normal`, or custom corporate styles).
   - Use `docx_inspect.py --mode styles` to discover valid styles available in the target document before appending text.

---

## 3. Skill Routing Matrix

Choose the appropriate skill based on user intent:

| User Intent | Appropriate Skill | Primary Script |
| :--- | :--- | :--- |
| Read/analyze document, convert to Markdown, extract images & links | `docx-to-markdown` | `scripts/docx_to_md.py` |
| Generate new `.docx` document from Markdown (100% Markdown syntax coverage) | `markdown-to-docx` | `scripts/md_to_docx.py` |
| Modify, append, replace text, or inspect existing `.docx` in-place | `docx-editor` | `scripts/docx_edit.py`, `scripts/docx_inspect.py` |
| Fine-tune formatting, apply corporate styles, or customize generated `.docx` | `docx-editor` | `scripts/docx_edit.py`, `scripts/docx_inspect.py` |
| Fill dynamic templates using Jinja2 syntax (loops, conditionals, tables) | `docx-editor` | `scripts/docx_render.py` |
| Convert `.docx` to `.pdf` (Word COM / LibreOffice headless) | `docx-to-pdf` | `scripts/docx_to_pdf.py` |

---

## 4. Separation of Concerns: Markdown Conversion vs. In-Place Editing & Styling

To ensure modularity and prevent bloated code, responsibilities between `markdown-to-docx` and `docx-editor` are strictly separated:

1. **`markdown-to-docx` Responsibilities (First-Stage Structural Conversion):**
   - Covers 100% of standard and extended Markdown syntax: headings (H1–H6), paragraphs, soft and hard line breaks, full inline formatting (bold, italic, strikethrough, inline code, nested combos), clickable hyperlinks, bulleted/numbered/nested lists, task lists/checkboxes (`- [ ]`, `- [x]`), GFM tables with column alignment, blockquotes and GitHub-style callouts/alerts, fenced code blocks with shading, embedded centered images, and horizontal rules.
   - It converts markdown into valid OpenXML Word documents with clean default styling or structural template style inheritance (`--template`).
   - **Non-goal:** It is **NOT** a desktop publishing or full OpenXML styling engine. Do not overload `markdown-to-docx` with complex Word-specific pagination, multi-section layout logic, table cell merges, or arbitrary corporate theme overrides.

2. **`docx-editor` Responsibilities (Second-Stage Fine-Tuning & Customization):**
   - Handles all fine-tuning, complex corporate styling, in-place text replacement, table row/column manipulation, and targeted OpenXML mutations on existing documents.
   - When a user requires custom corporate layouts, specific color overrides, in-place adjustments, or post-processing of a document originally generated from Markdown, the agent must **chain** the workflow:
     1. First, generate the `.docx` using `markdown-to-docx`.
     2. Second, inspect and fine-tune styles, replace text, or adjust tables using `docx-editor`.

