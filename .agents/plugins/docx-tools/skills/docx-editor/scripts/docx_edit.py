# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "python-docx>=1.1.2",
# ]
# ///
"""In-place safe editing of Microsoft Word (.docx) documents with style and layout preservation."""

import argparse
import copy
import os
import shutil
import sys
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def create_backup(docx_path: Path) -> Path:
    """Create a .bak copy of the target file before modification."""
    backup_path = docx_path.with_suffix(".docx.bak")
    try:
        shutil.copy2(docx_path, backup_path)
    except PermissionError:
        print(f"Error: Cannot create backup for '{docx_path.name}'. File may be locked by Word.", file=sys.stderr)
        sys.exit(1)
    return backup_path


def replace_text_in_paragraph(p: Paragraph, find_str: str, replace_str: str) -> int:
    """Replace all occurrences of find_str in paragraph across any run boundaries."""
    if not find_str or find_str not in p.text:
        return 0

    count = 0
    start_offset = 0

    while True:
        full_text = p.text
        idx = full_text.find(find_str, start_offset)
        if idx == -1:
            break

        match_end = idx + len(find_str)

        # Map character positions to runs
        cur_pos = 0
        start_run_idx = None
        start_run_char_offset = 0
        end_run_idx = None
        end_run_char_offset = 0

        runs = p.runs
        if not runs:
            break

        for r_idx, run in enumerate(runs):
            run_len = len(run.text)
            run_start = cur_pos
            run_end = cur_pos + run_len

            if start_run_idx is None and run_end > idx:
                start_run_idx = r_idx
                start_run_char_offset = idx - run_start

            if start_run_idx is not None and run_end >= match_end:
                end_run_idx = r_idx
                end_run_char_offset = match_end - run_start
                break

            cur_pos = run_end

        if start_run_idx is None or end_run_idx is None:
            break

        # Execute replacement
        if start_run_idx == end_run_idx:
            # Single run replacement
            r = runs[start_run_idx]
            r_text = r.text
            r.text = r_text[:start_run_char_offset] + replace_str + r_text[end_run_char_offset:]
        else:
            # Multi-run replacement
            start_run = runs[start_run_idx]
            end_run = runs[end_run_idx]

            start_prefix = start_run.text[:start_run_char_offset]
            end_suffix = end_run.text[end_run_char_offset:]

            start_run.text = start_prefix + replace_str
            end_run.text = end_suffix

            # Clear intermediate runs
            for mid_idx in range(start_run_idx + 1, end_run_idx):
                runs[mid_idx].text = ""

        count += 1
        start_offset = idx + len(replace_str)

    # Clean up empty runs from XML DOM
    for r in list(p.runs):
        if r.text == "":
            try:
                r._r.getparent().remove(r._r)
            except Exception:
                pass

    return count


def do_replace_text(doc: Document, find_str: str, replace_str: str) -> int:
    """Perform text replacement across body paragraphs, tables, headers, and footers."""
    total_replaced = 0

    # 1. Paragraphs in body
    for p in doc.paragraphs:
        total_replaced += replace_text_in_paragraph(p, find_str, replace_str)

    # 2. Paragraphs inside body tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    total_replaced += replace_text_in_paragraph(p, find_str, replace_str)

    # 3. Headers and Footers (paragraphs and tables) ONLY if defined in section OpenXML
    for sec in doc.sections:
        if sec._sectPr is not None and sec._sectPr.find(qn("w:headerReference")) is not None:
            for p in sec.header.paragraphs:
                total_replaced += replace_text_in_paragraph(p, find_str, replace_str)
            for table in sec.header.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            total_replaced += replace_text_in_paragraph(p, find_str, replace_str)

        if sec._sectPr is not None and sec._sectPr.find(qn("w:footerReference")) is not None:
            for p in sec.footer.paragraphs:
                total_replaced += replace_text_in_paragraph(p, find_str, replace_str)
            for table in sec.footer.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            total_replaced += replace_text_in_paragraph(p, find_str, replace_str)

    return total_replaced


def do_append_paragraph(doc: Document, text: str, style_name: str = None) -> Paragraph:
    """Append a paragraph to the end of the document, matching requested or default style."""
    if style_name and style_name in doc.styles:
        return doc.add_paragraph(text, style=style_name)
    elif doc.paragraphs:
        last_style = doc.paragraphs[-1].style.name
        try:
            return doc.add_paragraph(text, style=last_style)
        except Exception:
            return doc.add_paragraph(text)
    else:
        return doc.add_paragraph(text)


def do_insert_after(doc: Document, target_anchor: str, new_text: str, style_name: str = None) -> bool:
    """Insert a new paragraph immediately after an existing paragraph matching target_anchor."""
    for idx, p in enumerate(doc.paragraphs):
        if target_anchor.lower() in p.text.lower():
            new_p = doc.add_paragraph(new_text, style=style_name or p.style)
            p._element.addnext(new_p._element)
            return True
    return False


def do_add_table_row(doc: Document, table_index: int, values: list) -> bool:
    """Add a row to the specified table, inheriting formatting, borders, and styles from the previous row."""
    if table_index >= len(doc.tables):
        return False

    table = doc.tables[table_index]
    new_row = table.add_row()
    prev_row = table.rows[-2] if len(table.rows) > 1 else None

    # Copy row-level properties (trPr) like height and cantSplit, avoiding tblHeader
    if prev_row and prev_row._tr.trPr is not None:
        new_trPr = copy.deepcopy(prev_row._tr.trPr)
        for tblHeader in new_trPr.findall(qn("w:tblHeader")):
            new_trPr.remove(tblHeader)
        if new_row._tr.trPr is not None:
            new_row._tr.remove(new_row._tr.trPr)
        new_row._tr.insert(0, new_trPr)

    # Formatting tags to copy from source cell tcPr
    formatting_tags = {
        qn("w:tcBorders"),
        qn("w:shd"),
        qn("w:tcMar"),
        qn("w:vAlign"),
        qn("w:noWrap"),
        qn("w:textDirection"),
        qn("w:tcFitText"),
    }

    # Populate and style cells
    for idx, cell in enumerate(new_row.cells):
        if idx < len(values):
            cell.text = str(values[idx])
        else:
            cell.text = ""

        # Inherit cell & paragraph formatting from row above
        if prev_row and idx < len(prev_row.cells):
            src_cell = prev_row.cells[idx]

            # Copy cell XML properties (borders, background shading, margins, etc.)
            if src_cell._tc.tcPr is not None:
                tgt_tcPr = cell._tc.get_or_add_tcPr()
                for child in src_cell._tc.tcPr:
                    if child.tag in formatting_tags:
                        for existing in tgt_tcPr.findall(child.tag):
                            tgt_tcPr.remove(existing)
                        tgt_tcPr.append(copy.deepcopy(child))

            cell.vertical_alignment = src_cell.vertical_alignment

            if src_cell.paragraphs and cell.paragraphs:
                src_p = src_cell.paragraphs[0]
                tgt_p = cell.paragraphs[0]

                # Style inheritance
                try:
                    tgt_p.style = src_p.style
                except Exception:
                    pass

                tgt_p.alignment = src_p.alignment

                # Paragraph spacing inheritance
                try:
                    tgt_p.paragraph_format.space_before = src_p.paragraph_format.space_before
                    tgt_p.paragraph_format.space_after = src_p.paragraph_format.space_after
                    tgt_p.paragraph_format.line_spacing = src_p.paragraph_format.line_spacing
                except Exception:
                    pass

                # Font formatting inheritance
                if src_p.runs and tgt_p.runs:
                    src_run = src_p.runs[0]
                    tgt_run = tgt_p.runs[0]
                    tgt_run.font.name = src_run.font.name
                    tgt_run.font.size = src_run.font.size
                    tgt_run.font.bold = src_run.font.bold
                    tgt_run.font.italic = src_run.font.italic
                    tgt_run.font.underline = src_run.font.underline
                    if src_run.font.color and src_run.font.color.rgb:
                        tgt_run.font.color.rgb = src_run.font.color.rgb

    return True


def save_document_safely(doc: Document, target_path: Path):
    """Save document atomically using a temporary file to prevent corruption and handle locks."""
    temp_path = target_path.with_suffix(".tmp.docx")
    try:
        doc.save(str(temp_path))
        os.replace(temp_path, target_path)
    except PermissionError:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass
        print(
            f"Error: Permission denied saving '{target_path.name}'.\n"
            f"The file is likely locked by Microsoft Word or another application.\n"
            f"Please close Word and try again.",
            file=sys.stderr
        )
        sys.exit(1)
    except Exception as e:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass
        print(f"Error: Failed to save document '{target_path.name}': {e}", file=sys.stderr)
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description="Safely modify Microsoft Word (.docx) files without corrupting styles or layout."
    )
    parser.add_argument("docx_file", type=Path, help="Path to .docx document to modify")
    parser.add_argument("--no-backup", action="store_true", help="Do not create a .bak backup file")

    subparsers = parser.add_subparsers(dest="action", required=True)

    # Subcommand: replace-text
    p_replace = subparsers.add_parser("replace-text", help="Find and replace text across document")
    p_replace.add_argument("--find", required=True, help="Text to search for")
    p_replace.add_argument("--replace", required=True, help="Replacement text")

    # Subcommand: append-paragraph
    p_append = subparsers.add_parser("append-paragraph", help="Append paragraph to end of document")
    p_append.add_argument("--text", required=True, help="Paragraph text to append")
    p_append.add_argument("--style", default=None, help="Existing Word style name (e.g. 'Normal', 'Heading 2')")

    # Subcommand: insert-after
    p_insert = subparsers.add_parser("insert-after", help="Insert paragraph after a specific anchor text")
    p_insert.add_argument("--anchor", required=True, help="Text or heading to find as anchor")
    p_insert.add_argument("--text", required=True, help="New paragraph text to insert")
    p_insert.add_argument("--style", default=None, help="Word style name (defaults to anchor paragraph's style)")

    # Subcommand: add-table-row
    p_row = subparsers.add_parser("add-table-row", help="Append a row to a table")
    p_row.add_argument("--table-index", type=int, default=0, help="0-based index of the table")
    p_row.add_argument("--values", nargs="+", required=True, help="Values for table row cells")

    args = parser.parse_args()

    if not args.docx_file.is_file():
        print(f"Error: File not found: {args.docx_file}", file=sys.stderr)
        sys.exit(1)

    docx_path = args.docx_file.resolve()

    # Backup
    if not args.no_backup:
        bak_file = create_backup(docx_path)
        print(f"Backup created: {bak_file.name}")

    try:
        doc = Document(str(docx_path))
    except PermissionError:
        print(f"Error: Permission denied opening '{docx_path.name}'. File may be locked by Word.", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: Failed to open '{docx_path.name}': {e}", file=sys.stderr)
        sys.exit(1)

    success = False

    if args.action == "replace-text":
        count = do_replace_text(doc, args.find, args.replace)
        print(f"Text replacement complete: {count} occurrence(s) replaced.")
        success = count > 0

    elif args.action == "append-paragraph":
        do_append_paragraph(doc, args.text, args.style)
        print(f"Appended new paragraph to '{docx_path.name}'.")
        success = True

    elif args.action == "insert-after":
        ok = do_insert_after(doc, args.anchor, args.text, args.style)
        if ok:
            print(f"Successfully inserted paragraph after anchor '{args.anchor}'.")
            success = True
        else:
            print(f"Warning: Anchor text '{args.anchor}' was not found in document.", file=sys.stderr)
            sys.exit(1)

    elif args.action == "add-table-row":
        ok = do_add_table_row(doc, args.table_index, args.values)
        if ok:
            print(f"Successfully added row to table #{args.table_index}.")
            success = True
        else:
            print(f"Error: Table #{args.table_index} does not exist in document.", file=sys.stderr)
            sys.exit(1)

    if success:
        save_document_safely(doc, docx_path)
        print(f"SUCCESS: Saved changes to '{docx_path.name}' while preserving all styles and formatting.")


if __name__ == "__main__":
    main()
