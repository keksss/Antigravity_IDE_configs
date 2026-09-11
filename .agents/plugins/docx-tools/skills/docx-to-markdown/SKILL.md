---
name: docx-to-markdown
description: >-
  Converts Microsoft Word (.docx) documents into clean Markdown files with full extraction
  of embedded images into a dedicated media folder, relative image links, hyperlinks, and tables.
  Use when the user asks to "convert docx to markdown", "extract text and images from word",
  "unpack docx", or import word documentation into markdown.
---

# docx-to-markdown Skill

Convert `.docx` documents into structured Markdown while extracting all embedded images into a media folder and preserving hyperlinks, lists, and tables.

---

## Environment Setup

This skill uses Python with `mammoth` and `python-docx`. Dependencies are defined via PEP 723 inline metadata in the script header.

Execute scripts using either:
1. **With `uv` (Recommended):**
   ```bash
   uv run scripts/docx_to_md.py <path_to_docx>
   ```
2. **With Python and local virtual environment:**
   If `.venv` does not exist in this skill folder, create it:
   ```bash
   # Windows
   python -m venv .venv
   .\.venv\Scripts\pip install python-docx mammoth
   .\.venv\Scripts\python scripts/docx_to_md.py <path_to_docx>

   # Linux/macOS
   python -m venv .venv
   ./.venv/bin/pip install python-docx mammoth
   ./.venv/bin/python scripts/docx_to_md.py <path_to_docx>
   ```

---

## Usage Scenarios & Commands

### 1. Default Conversion (Generates `.md` and `<filename>_media/` next to document)
```bash
uv run scripts/docx_to_md.py document.docx
```
- Produces: `document.md`
- Creates folder: `document_media/` containing `image_001.png`, `image_002.jpg`, etc.
- In `document.md`, images are linked using relative paths: `![Image 1](document_media/image_001.png)`

### 2. Custom Output File and Media Folder
```bash
uv run scripts/docx_to_md.py path/to/report.docx -o output/report.md --media-dir output/assets/images
```

---

## CLI Options Reference

| Argument | Description | Default |
| :--- | :--- | :--- |
| `docx_file` | Path to the `.docx` document to convert. | *(Required)* |
| `-o`, `--output` | Destination path for the `.md` file. | `<docx_name>.md` |
| `--media-dir` | Directory to store extracted images. | `<docx_name>_media` |
| `--engine` | Conversion engine: `auto`, `markitdown` (Microsoft), or `mammoth`. | `auto` |
| `--include-headers` | Include header and footer text at the top of the Markdown output. | `false` |
| `--no-comments` | Disable automatic extraction of Word comments into Markdown. | `false` (comments enabled) |

---

## Output Behavior

1. **Images:** Extracted as original files into `--media-dir` with clean relative links in the Markdown.
2. **Comments:** Extracted from OpenXML (`word/comments.xml`), rendered in a dedicated `## Комментарии / Comments` section with author, date, target fragment quote, and comment text.
3. **Tables:** Rendered as standard GFM Markdown table rows with headers and delimiters.
4. **Links:** Word hyperlinks (`w:hyperlink`) are converted to `[Text](https://...)`.
5. **Headings:** Preserves levels H1 through H6 based on Word heading styles (including localized Russian styles).

