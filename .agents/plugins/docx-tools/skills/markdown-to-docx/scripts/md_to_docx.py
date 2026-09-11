# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "python-docx>=1.1.2",
#     "markdown-it-py>=3.0.0",
#     "Pillow>=10.0.0",
# ]
# ///
"""Convert Markdown (.md) documents to Microsoft Word (.docx) with 100% syntax coverage and optional template styling."""

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path
from PIL import Image
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from markdown_it import MarkdownIt
from markdown_it.tree import SyntaxTreeNode

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def clear_document_body(doc: Document):
    """Remove existing paragraphs and tables from a template while preserving styles, headers, footers, and sections."""
    body_elm = doc._body._element
    for child in list(body_elm):
        if child.tag.endswith("sectPr"):
            continue
        body_elm.remove(child)


def get_node_info(node):
    """Safely extract node attributes whether node is a SyntaxTreeNode or a dictionary."""
    if isinstance(node, dict):
        return node.get("type", ""), node.get("content", ""), node.get("children", []), node.get("attrs", {})
    return getattr(node, "type", ""), getattr(node, "content", ""), getattr(node, "children", []), getattr(node, "attrs", {})


def extract_plain_text(node) -> str:
    """Recursively extract plain text from any AST node."""
    node_type, content, children, _ = get_node_info(node)
    if content:
        return content
    texts = []
    for c in children:
        texts.append(extract_plain_text(c))
    return "".join(texts)


def add_image_smart(doc: Document, paragraph, img_path: Path, max_w_in: float = 5.8, max_h_in: float = 7.5):
    """Scale image preserving natural dimensions, aspect ratio, and DPI via Pillow."""
    if not img_path.is_file():
        r = paragraph.add_run(f"[Missing Image: {img_path.name}]")
        r.italic = True
        r.font.color.rgb = RGBColor(180, 50, 50)
        return

    try:
        with Image.open(img_path) as im:
            px_w, px_h = im.size
            dpi = im.info.get("dpi", (96, 96))
            dpi_x = dpi[0] if isinstance(dpi, tuple) and dpi[0] > 0 else 96

        nat_w = px_w / dpi_x
        nat_h = px_h / dpi_x

        scale = 1.0
        if nat_w > max_w_in:
            scale = min(scale, max_w_in / nat_w)
        if nat_h > max_h_in:
            scale = min(scale, max_h_in / nat_h)

        final_w = Inches(nat_w * scale)

        pic_p = doc.add_paragraph()
        pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic_p.paragraph_format.space_before = Pt(4)
        pic_p.paragraph_format.space_after = Pt(4)
        pic_p.add_run().add_picture(str(img_path), width=final_w)
    except Exception as e:
        r = paragraph.add_run(f"[Image: {img_path.name} (Error: {e})]")
        r.italic = True


def render_inline_tree(
    paragraph,
    nodes: list,
    base_dir: Path,
    doc: Document,
    bold: bool = False,
    italic: bool = False,
    strike: bool = False,
    code: bool = False,
    hyperlink_el=None,
):
    """Recursively render inline Markdown AST nodes with full stacked formatting (bold, italic, strike, code, links)."""
    for node in nodes:
        node_type, content, children, attrs = get_node_info(node)

        if node_type == "text":
            text = content
            if not text:
                continue

            if hyperlink_el is not None:
                r = OxmlElement("w:r")
                rPr = OxmlElement("w:rPr")
                if bold:
                    rPr.append(OxmlElement("w:b"))
                if italic:
                    rPr.append(OxmlElement("w:i"))
                if strike:
                    rPr.append(OxmlElement("w:strike"))
                if code:
                    rFonts = OxmlElement("w:rFonts")
                    rFonts.set(qn("w:ascii"), "Consolas")
                    rFonts.set(qn("w:hAnsi"), "Consolas")
                    rPr.append(rFonts)
                    sz = OxmlElement("w:sz")
                    sz.set(qn("w:val"), "19")
                    rPr.append(sz)

                # Default link colors and single underline
                c = OxmlElement("w:color")
                c.set(qn("w:val"), "0563C1")
                rPr.append(c)
                u = OxmlElement("w:u")
                u.set(qn("w:val"), "single")
                rPr.append(u)

                r.append(rPr)
                t = OxmlElement("w:t")
                t.text = text
                if text.startswith(" ") or text.endswith(" "):
                    t.set(qn("xml:space"), "preserve")
                r.append(t)
                hyperlink_el.append(r)
            else:
                run = paragraph.add_run(text)
                run.bold = bold
                run.italic = italic
                run.font.strike = strike
                if code:
                    run.font.name = "Consolas"
                    run.font.size = Pt(9.5)
                    run.font.color.rgb = RGBColor(199, 37, 78)

        elif node_type == "code_inline":
            if hyperlink_el is not None:
                r = OxmlElement("w:r")
                rPr = OxmlElement("w:rPr")
                if bold:
                    rPr.append(OxmlElement("w:b"))
                if italic:
                    rPr.append(OxmlElement("w:i"))
                rFonts = OxmlElement("w:rFonts")
                rFonts.set(qn("w:ascii"), "Consolas")
                rFonts.set(qn("w:hAnsi"), "Consolas")
                rPr.append(rFonts)
                sz = OxmlElement("w:sz")
                sz.set(qn("w:val"), "19")
                rPr.append(sz)
                c = OxmlElement("w:color")
                c.set(qn("w:val"), "C7254E")
                rPr.append(c)
                r.append(rPr)
                t = OxmlElement("w:t")
                t.text = content
                if content.startswith(" ") or content.endswith(" "):
                    t.set(qn("xml:space"), "preserve")
                r.append(t)
                hyperlink_el.append(r)
            else:
                run = paragraph.add_run(content)
                run.bold = bold
                run.italic = italic
                run.font.strike = strike
                run.font.name = "Consolas"
                run.font.size = Pt(9.5)
                run.font.color.rgb = RGBColor(199, 37, 78)

        elif node_type == "strong":
            if children:
                render_inline_tree(paragraph, children, base_dir, doc, bold=True, italic=italic, strike=strike, code=code, hyperlink_el=hyperlink_el)
            elif content:
                render_inline_tree(
                    paragraph,
                    [{"type": "text", "content": content, "children": [], "attrs": {}}],
                    base_dir,
                    doc,
                    bold=True,
                    italic=italic,
                    strike=strike,
                    code=code,
                    hyperlink_el=hyperlink_el,
                )

        elif node_type == "em":
            if children:
                render_inline_tree(paragraph, children, base_dir, doc, bold=bold, italic=True, strike=strike, code=code, hyperlink_el=hyperlink_el)
            elif content:
                render_inline_tree(
                    paragraph,
                    [{"type": "text", "content": content, "children": [], "attrs": {}}],
                    base_dir,
                    doc,
                    bold=bold,
                    italic=True,
                    strike=strike,
                    code=code,
                    hyperlink_el=hyperlink_el,
                )

        elif node_type == "s":
            if children:
                render_inline_tree(paragraph, children, base_dir, doc, bold=bold, italic=italic, strike=True, code=code, hyperlink_el=hyperlink_el)
            elif content:
                render_inline_tree(
                    paragraph,
                    [{"type": "text", "content": content, "children": [], "attrs": {}}],
                    base_dir,
                    doc,
                    bold=bold,
                    italic=italic,
                    strike=True,
                    code=code,
                    hyperlink_el=hyperlink_el,
                )

        elif node_type == "softbreak":
            if hyperlink_el is not None:
                r = OxmlElement("w:r")
                t = OxmlElement("w:t")
                t.text = " "
                t.set(qn("xml:space"), "preserve")
                r.append(t)
                hyperlink_el.append(r)
            else:
                paragraph.add_run(" ")

        elif node_type == "hardbreak":
            if hyperlink_el is not None:
                r = OxmlElement("w:r")
                r.append(OxmlElement("w:br"))
                hyperlink_el.append(r)
            else:
                paragraph.add_run().add_break()

        elif node_type == "link":
            href = attrs.get("href", "")
            try:
                part = paragraph.part
                r_id = part.relate_to(href, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
                hyperlink_obj = OxmlElement("w:hyperlink")
                hyperlink_obj.set(qn("r:id"), r_id)

                if children:
                    render_inline_tree(
                        paragraph,
                        children,
                        base_dir,
                        doc,
                        bold=bold,
                        italic=italic,
                        strike=strike,
                        code=code,
                        hyperlink_el=hyperlink_obj,
                    )
                else:
                    r = OxmlElement("w:r")
                    rPr = OxmlElement("w:rPr")
                    c = OxmlElement("w:color")
                    c.set(qn("w:val"), "0563C1")
                    rPr.append(c)
                    u = OxmlElement("w:u")
                    u.set(qn("w:val"), "single")
                    rPr.append(u)
                    r.append(rPr)
                    t = OxmlElement("w:t")
                    t.text = href
                    r.append(t)
                    hyperlink_obj.append(r)

                paragraph._p.append(hyperlink_obj)
            except Exception:
                link_text = extract_plain_text(node) or href
                run = paragraph.add_run(f"{link_text} ({href})")
                run.underline = True
                run.font.color.rgb = RGBColor(5, 99, 193)

        elif node_type == "image":
            src = attrs.get("src", "")
            img_path = (base_dir / src).resolve() if not os.path.isabs(src) else Path(src)
            add_image_smart(doc, paragraph, img_path)

        elif node_type in ("paragraph", "inline", "list_item"):
            if children:
                render_inline_tree(paragraph, children, base_dir, doc, bold=bold, italic=italic, strike=strike, code=code, hyperlink_el=hyperlink_el)
            elif content:
                run = paragraph.add_run(content)
                run.bold = bold
                run.italic = italic

        elif children:
            render_inline_tree(paragraph, children, base_dir, doc, bold=bold, italic=italic, strike=strike, code=code, hyperlink_el=hyperlink_el)
        elif content:
            run = paragraph.add_run(content)
            run.bold = bold
            run.italic = italic


def add_code_block(doc: Document, code_text: str, language: str = ""):
    """Render code block with background shading, left accent border, and Consolas font."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)

    # Shading (light gray #F5F6F8)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "F4F5F7")
    pPr.append(shd)

    # Left accent border (GitHub blue #0969DA)
    pBdr = OxmlElement("w:pBdr")
    left_bdr = OxmlElement("w:left")
    left_bdr.set(qn("w:val"), "single")
    left_bdr.set(qn("w:sz"), "18")  # 2.25 pt
    left_bdr.set(qn("w:space"), "8")
    left_bdr.set(qn("w:color"), "0969DA")
    pBdr.append(left_bdr)
    pPr.append(pBdr)

    clean_code = code_text.rstrip("\n")
    run = p.add_run(clean_code)
    run.font.name = "Consolas"
    run.font.size = Pt(9.5)
    run.font.color.rgb = RGBColor(36, 41, 46)


def strip_alert_tag(inline_nodes: list, alert_type: str) -> list:
    """Strip [!ALERT] prefix from the first inline text node without mutating AST in-place."""
    if not inline_nodes:
        return inline_nodes
    tag_str = f"[!{alert_type}]"
    first = inline_nodes[0]
    n_type, content, children, attrs = get_node_info(first)
    if n_type == "text" and content:
        cleaned = content.lstrip()
        if cleaned.startswith(tag_str):
            new_content = cleaned[len(tag_str):].lstrip("\n ")
            new_first = {"type": "text", "content": new_content, "children": children, "attrs": attrs}
            return [new_first] + inline_nodes[1:]
    return inline_nodes


def render_callout_block(node: SyntaxTreeNode, alert_type: str, doc: Document, base_dir: Path):
    """Render GitHub-style alert callouts (> [!NOTE], etc.) with themed accent borders and full inline Markdown."""
    color_map = {
        "NOTE": ("0969DA", "Примечание"),
        "TIP": ("1A7F37", "Совет"),
        "IMPORTANT": ("8250DF", "Важно"),
        "WARNING": ("9A6700", "Внимание"),
        "CAUTION": ("D1242F", "Осторожно"),
    }
    hex_color, label = color_map.get(alert_type.upper(), ("0969DA", alert_type))

    paragraphs = [c for c in node.children if c.type == "paragraph"]
    if not paragraphs:
        paragraphs = [node]

    for p_idx, p_node in enumerate(paragraphs):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_before = Pt(4) if p_idx == 0 else Pt(2)
        p.paragraph_format.space_after = Pt(6) if p_idx == len(paragraphs) - 1 else Pt(2)

        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "F8FAFD")
        pPr.append(shd)

        pBdr = OxmlElement("w:pBdr")
        left_bdr = OxmlElement("w:left")
        left_bdr.set(qn("w:val"), "single")
        left_bdr.set(qn("w:sz"), "24")
        left_bdr.set(qn("w:space"), "8")
        left_bdr.set(qn("w:color"), hex_color)
        pBdr.append(left_bdr)
        pPr.append(pBdr)

        if p_idx == 0:
            # Add alert header badge
            lbl_run = p.add_run(f"[{label}] ")
            lbl_run.bold = True
            r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
            lbl_run.font.color.rgb = RGBColor(r, g, b)

        # Extract and render inline children
        inline_nodes = []
        for c in p_node.children:
            if c.type == "inline" and c.children:
                inline_nodes.extend(c.children)
            else:
                inline_nodes.append(c)

        if p_idx == 0:
            inline_nodes = strip_alert_tag(inline_nodes, alert_type)

        render_inline_tree(p, inline_nodes, base_dir, doc)


def render_regular_blockquote(node: SyntaxTreeNode, doc: Document, base_dir: Path):
    """Render standard blockquote preserving multiple paragraphs and quotes."""
    paragraphs = [c for c in node.children if c.type == "paragraph"]
    if not paragraphs:
        paragraphs = [node]

    for p_node in paragraphs:
        try:
            p = doc.add_paragraph(style="Quote")
        except Exception:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.25)

        # Left subtle gray border
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        left_bdr = OxmlElement("w:left")
        left_bdr.set(qn("w:val"), "single")
        left_bdr.set(qn("w:sz"), "18")
        left_bdr.set(qn("w:space"), "8")
        left_bdr.set(qn("w:color"), "D0D7DE")
        pBdr.append(left_bdr)
        pPr.append(pBdr)

        inline_nodes = []
        for c in p_node.children:
            if c.type == "inline" and c.children:
                inline_nodes.extend(c.children)
            else:
                inline_nodes.append(c)

        render_inline_tree(p, inline_nodes, base_dir, doc, italic=True)


def render_table(node: SyntaxTreeNode, doc: Document, base_dir: Path):
    """Render markdown table with column alignments, rich cell formatting, and repeated header rows."""
    rows_nodes = []
    for section in node.children:  # thead, tbody
        for tr in section.children:
            row_cells = []
            for th_or_td in tr.children:
                style_str = th_or_td.attrs.get("style", "")
                align_str = th_or_td.attrs.get("align", "")
                align = "left"
                if "center" in style_str or align_str == "center":
                    align = "center"
                elif "right" in style_str or align_str == "right":
                    align = "right"
                row_cells.append((th_or_td, align))
            rows_nodes.append((row_cells, section.type == "thead"))

    if not rows_nodes:
        return

    num_rows = len(rows_nodes)
    num_cols = max(len(cells) for cells, _ in rows_nodes)

    table = doc.add_table(rows=num_rows, cols=num_cols)
    try:
        table.style = "Table Grid"
    except Exception:
        pass

    for r_idx, (cells, is_header) in enumerate(rows_nodes):
        row = table.rows[r_idx]
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        if is_header:
            trPr.append(OxmlElement("w:tblHeader"))

        for c_idx, (th_or_td, align) in enumerate(cells):
            if c_idx < num_cols:
                cell = table.cell(r_idx, c_idx)
                p = cell.paragraphs[0]
                p.text = ""
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)

                if align == "center":
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif align == "right":
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                if is_header:
                    tcPr = cell._tc.get_or_add_tcPr()
                    shd = OxmlElement("w:shd")
                    shd.set(qn("w:val"), "clear")
                    shd.set(qn("w:color"), "auto")
                    shd.set(qn("w:fill"), "F2F4F7")
                    tcPr.append(shd)

                inline_nodes = []
                if th_or_td.children:
                    for c in th_or_td.children:
                        if c.type == "inline" and c.children:
                            inline_nodes.extend(c.children)
                        else:
                            inline_nodes.append(c)
                elif th_or_td.content:
                    inline_nodes = [{"type": "text", "content": th_or_td.content.strip(), "children": [], "attrs": {}}]

                render_inline_tree(p, inline_nodes, base_dir, doc, bold=is_header)

    # Empty paragraph after table for standard spacing
    doc.add_paragraph()


def transform_task_list(inline_nodes: list) -> list:
    """Convert [ ] and [x] markdown task list markers into Unicode checkbox glyphs."""
    if not inline_nodes:
        return inline_nodes
    first = inline_nodes[0]
    n_type, content, children, attrs = get_node_info(first)
    if n_type == "text" and content:
        if content.startswith("[ ] "):
            new_first = {"type": "text", "content": "☐ " + content[4:], "children": children, "attrs": attrs}
            return [new_first] + inline_nodes[1:]
        elif content.startswith("[x] ") or content.startswith("[X] "):
            new_first = {"type": "text", "content": "☑ " + content[4:], "children": children, "attrs": attrs}
            return [new_first] + inline_nodes[1:]
        elif content == "[ ]":
            new_first = {"type": "text", "content": "☐ ", "children": children, "attrs": attrs}
            return [new_first] + inline_nodes[1:]
        elif content in ("[x]", "[X]"):
            new_first = {"type": "text", "content": "☑ ", "children": children, "attrs": attrs}
            return [new_first] + inline_nodes[1:]
    return inline_nodes


def render_list_items(node: SyntaxTreeNode, doc: Document, base_dir: Path, is_ordered: bool, level: int = 1):
    """Recursively render nested bullet and ordered lists preserving hierarchy and task list checkboxes."""
    for item in node.children:
        style_prefix = "List Number" if is_ordered else "List Bullet"
        style_name = style_prefix if level == 1 else f"{style_prefix} {min(level, 3)}"

        inline_children = []
        nested_lists = []

        for c in item.children:
            if c.type in ("bullet_list", "ordered_list"):
                nested_lists.append(c)
            elif c.type == "paragraph":
                for sub in c.children:
                    if sub.type == "inline" and sub.children:
                        inline_children.extend(sub.children)
                    else:
                        inline_children.append(sub)
            elif c.type == "inline" and c.children:
                inline_children.extend(c.children)
            else:
                inline_children.append(c)

        # Check and apply task list checkboxes
        inline_children = transform_task_list(inline_children)

        try:
            p = doc.add_paragraph(style=style_name)
        except Exception:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.25 * level)

        render_inline_tree(p, inline_children, base_dir, doc)

        for nl in nested_lists:
            render_list_items(nl, doc, base_dir, is_ordered=(nl.type == "ordered_list"), level=level + 1)


def convert_with_pandoc(md_path: Path, output_path: Path, template_path: Path = None) -> bool:
    """High-fidelity conversion via Pandoc binary if available."""
    pandoc_bin = shutil.which("pandoc")
    if not pandoc_bin:
        return False

    cmd = [pandoc_bin, str(md_path), "-f", "gfm", "-t", "docx", "-o", str(output_path)]
    if template_path and template_path.is_file():
        cmd.extend(["--reference-doc", str(template_path)])

    try:
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=60)
        return res.returncode == 0 and output_path.is_file() and output_path.stat().st_size > 0
    except Exception:
        return False


def convert_md_to_docx(md_path: Path, output_path: Path, template_path: Path = None, engine: str = "auto") -> dict:
    """Convert a Markdown file to DOCX with 100% syntax coverage and optional reference template."""
    # 1. Try Pandoc engine if explicitly requested or auto
    if engine in ("auto", "pandoc"):
        if convert_with_pandoc(md_path, output_path, template_path):
            doc_test = Document(str(output_path))
            return {
                "engine": "pandoc",
                "md_path": str(md_path),
                "output_path": str(output_path),
                "template_used": str(template_path) if template_path else None,
                "paragraphs_generated": len(doc_test.paragraphs),
                "tables_generated": len(doc_test.tables),
            }
        elif engine == "pandoc":
            raise RuntimeError("Pandoc engine was requested, but pandoc executable was not found.")

    # 2. Enhanced Python Engine
    with open(md_path, "r", encoding="utf-8") as f:
        md_content = f.read()

    base_dir = md_path.parent

    if template_path and template_path.is_file():
        doc = Document(str(template_path))
        clear_document_body(doc)
    else:
        doc = Document()
        style_normal = doc.styles["Normal"]
        font = style_normal.font
        font.name = "Calibri"
        font.size = Pt(11)

    md = MarkdownIt("commonmark").enable("table").enable("strikethrough")
    tokens = md.parse(md_content)
    root = SyntaxTreeNode(tokens)

    for node in root.children:
        if node.type == "heading":
            level = int(node.tag[1]) if len(node.tag) > 1 and node.tag[1].isdigit() else 1
            heading_style = f"Heading {min(max(level, 1), 6)}"
            try:
                p = doc.add_paragraph(style=heading_style)
            except Exception:
                try:
                    p = doc.add_paragraph(style=f"Заголовок {min(max(level, 1), 6)}")
                except Exception:
                    p = doc.add_paragraph()
                    p.style = "Normal"

            inline_nodes = []
            for c in node.children:
                if c.type == "inline" and c.children:
                    inline_nodes.extend(c.children)
                else:
                    inline_nodes.append(c)

            if inline_nodes:
                render_inline_tree(p, inline_nodes, base_dir, doc)
            elif node.content:
                p.add_run(node.content)

        elif node.type == "paragraph":
            p = doc.add_paragraph(style="Normal")
            inline_nodes = []
            for c in node.children:
                if c.type == "inline" and c.children:
                    inline_nodes.extend(c.children)
                else:
                    inline_nodes.append(c)

            if inline_nodes:
                render_inline_tree(p, inline_nodes, base_dir, doc)
            elif node.content:
                p.add_run(node.content)

        elif node.type == "bullet_list":
            render_list_items(node, doc, base_dir, is_ordered=False, level=1)

        elif node.type == "ordered_list":
            render_list_items(node, doc, base_dir, is_ordered=True, level=1)

        elif node.type == "blockquote":
            # Check for GitHub-style alerts: > [!NOTE], > [!WARNING], etc.
            plain = extract_plain_text(node).strip()
            is_alert = False
            for alert_name in ["NOTE", "TIP", "IMPORTANT", "WARNING", "CAUTION"]:
                tag = f"[!{alert_name}]"
                if plain.startswith(tag):
                    render_callout_block(node, alert_name, doc, base_dir)
                    is_alert = True
                    break

            if not is_alert:
                render_regular_blockquote(node, doc, base_dir)

        elif node.type in ("fence", "code_block"):
            lang = getattr(node, "info", "").strip()
            add_code_block(doc, node.content, language=lang)

        elif node.type == "table":
            render_table(node, doc, base_dir)

        elif node.type == "hr":
            p = doc.add_paragraph("―" * 35)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(8)
            for r in p.runs:
                r.font.color.rgb = RGBColor(160, 160, 160)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(".tmp.docx")
    doc.save(str(temp_path))
    os.replace(temp_path, output_path)

    return {
        "engine": "python",
        "md_path": str(md_path),
        "output_path": str(output_path),
        "template_used": str(template_path) if template_path else None,
        "paragraphs_generated": len(doc.paragraphs),
        "tables_generated": len(doc.tables),
    }


def main():
    parser = argparse.ArgumentParser(
        description="Convert Markdown (.md) to a professionally formatted Word (.docx) document."
    )
    parser.add_argument("md_file", type=Path, help="Path to input Markdown (.md) file")
    parser.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Path for destination .docx document (defaults to <name>.docx)"
    )
    parser.add_argument(
        "-t", "--template", type=Path, default=None,
        help="Optional reference .docx template (inherits styles, fonts, margins, headers/footers)"
    )
    parser.add_argument(
        "--engine", choices=["auto", "python", "pandoc"], default="auto",
        help="Conversion engine to use (default: auto)"
    )

    args = parser.parse_args()

    if not args.md_file.is_file():
        print(f"Error: Markdown file not found: {args.md_file}", file=sys.stderr)
        sys.exit(1)

    md_path = args.md_file.resolve()
    output_docx = args.output.resolve() if args.output else md_path.parent / f"{md_path.stem}.docx"
    template_path = args.template.resolve() if args.template else None

    if template_path and not template_path.is_file():
        print(f"Warning: Template file '{template_path}' does not exist. Proceeding with clean default styling.", file=sys.stderr)
        template_path = None

    res = convert_md_to_docx(md_path, output_docx, template_path, engine=args.engine)

    print(f"SUCCESS: Generated '{output_docx.name}' (engine: {res['engine']}).")
    print(f"  Destination: {output_docx}")
    if res["template_used"]:
        print(f"  Template:    {res['template_used']} (inherited custom styles)")
    else:
        print(f"  Template:    None (clean default formatting)")
    print(f"  Elements:    {res['paragraphs_generated']} paragraphs, {res['tables_generated']} tables")


if __name__ == "__main__":
    main()
