# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "docxtpl>=0.16.8",
#     "python-docx>=1.1.2",
#     "jinja2>=3.1.0",
# ]
# ///
"""Generate Microsoft Word (.docx) documents from Jinja2-enabled templates using docxtpl."""

import argparse
import json
import os
import shutil
import sys
import re
from pathlib import Path
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm, Inches

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


class RobustDocxTemplate(DocxTemplate):
    """Subclass of DocxTemplate that normalizes table and paragraph tags."""
    def patch_xml(self, src_xml: str) -> str:
        # Normalize tags with spaces like '{% tr ' to '{%tr '
        src_xml = re.sub(r"{%\s*(tr|tc|p|r)\s+", r"{%\1 ", src_xml)
        # Normalize '{% trendfor %}' and '{% tr endfor %}' to '{%tr endfor %}'
        src_xml = re.sub(r"{%\s*(?:trendfor|tr\s*endfor)\s*%}", "{%tr endfor %}", src_xml)
        return super().patch_xml(src_xml)


def create_backup(docx_path: Path) -> Path:
    """Create a .bak copy of the target file if it already exists."""
    backup_path = docx_path.with_suffix(".docx.bak")
    try:
        shutil.copy2(docx_path, backup_path)
    except PermissionError:
        print(f"Error: Cannot create backup for '{docx_path.name}'. File may be locked by Word.", file=sys.stderr)
        sys.exit(1)
    return backup_path


def process_context_images(context: dict, doc: DocxTemplate, base_dir: Path) -> dict:
    """Recursively convert image dicts like {'__image__': 'path/img.png', 'width_mm': 50} to InlineImage objects."""
    processed = {}
    for k, v in context.items():
        if isinstance(v, dict):
            if "__image__" in v:
                img_path = Path(v["__image__"])
                if not img_path.is_absolute():
                    img_path = base_dir / img_path
                if img_path.is_file():
                    width = Mm(v.get("width_mm", 40)) if "width_mm" in v else (Inches(v.get("width_in")) if "width_in" in v else None)
                    height = Mm(v.get("height_mm", 40)) if "height_mm" in v else (Inches(v.get("height_in")) if "height_in" in v else None)
                    processed[k] = InlineImage(doc, str(img_path), width=width, height=height)
                else:
                    processed[k] = f"[Image Not Found: {v['__image__']}]"
            else:
                processed[k] = process_context_images(v, doc, base_dir)
        elif isinstance(v, list):
            processed[k] = [
                process_context_images(item, doc, base_dir) if isinstance(item, dict) else item
                for item in v
            ]
        else:
            processed[k] = v
    return processed


def render_template(template_path: Path, output_path: Path, context: dict, base_dir: Path) -> dict:
    """Render a docxtpl template with the provided context dictionary."""
    if not template_path.is_file():
        raise FileNotFoundError(f"Template not found: {template_path}")

    doc = RobustDocxTemplate(str(template_path))
    processed_context = process_context_images(context, doc, base_dir)

    doc.render(processed_context)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(".tmp.docx")

    try:
        doc.save(str(temp_path))
        os.replace(temp_path, output_path)
    except PermissionError:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass
        print(
            f"Error: Permission denied saving '{output_path.name}'.\n"
            f"The file may be open in Microsoft Word. Please close Word and retry.",
            file=sys.stderr
        )
        sys.exit(1)

    return {
        "template": str(template_path),
        "output": str(output_path),
        "variables_rendered": list(context.keys()),
    }


def inspect_template_variables(template_path: Path) -> list:
    """Extract all undeclared Jinja2 template variables from a docx template."""
    doc = RobustDocxTemplate(str(template_path))
    variables = sorted(list(doc.get_undeclared_template_variables()))
    return variables


def main():
    parser = argparse.ArgumentParser(
        description="Fill Microsoft Word (.docx) templates using Jinja2 syntax via docxtpl."
    )
    parser.add_argument("template", type=Path, help="Path to input template .docx file")
    parser.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Path for destination .docx document"
    )
    parser.add_argument(
        "-d", "--data", type=Path, default=None,
        help="Path to JSON file with context data"
    )
    parser.add_argument(
        "--json", type=str, default=None,
        help="Inline JSON string with context data"
    )
    parser.add_argument(
        "--inspect-vars", action="store_true",
        help="Print all Jinja2 variables found in the template and exit"
    )
    parser.add_argument(
        "--no-backup", action="store_true",
        help="Do not create a .bak backup if output file already exists"
    )

    args = parser.parse_args()

    template_path = args.template.resolve()
    if not template_path.is_file():
        print(f"Error: Template not found: {template_path}", file=sys.stderr)
        sys.exit(1)

    # Variable inspection mode
    if args.inspect_vars:
        vars_found = inspect_template_variables(template_path)
        print(f"=== Template Variables in '{template_path.name}' ({len(vars_found)}) ===")
        if vars_found:
            for v in vars_found:
                print(f"  - {{{{ {v} }}}}")
        else:
            print("  No Jinja2 variables found in template.")
        return

    # Normal render mode
    if not args.output:
        print("Error: Output path is required for rendering. Use -o / --output.", file=sys.stderr)
        sys.exit(1)

    context = {}
    if args.data:
        data_path = args.data.resolve()
        if not data_path.is_file():
            print(f"Error: Data file not found: {data_path}", file=sys.stderr)
            sys.exit(1)
        with open(data_path, "r", encoding="utf-8") as f:
            context = json.load(f)
    elif args.json:
        try:
            context = json.loads(args.json)
        except json.JSONDecodeError as e:
            print(f"Error: Invalid inline JSON: {e}", file=sys.stderr)
            sys.exit(1)
    else:
        print("Warning: No context data provided (--data or --json). Rendering with empty context.", file=sys.stderr)

    output_path = args.output.resolve()
    base_dir = args.data.resolve().parent if args.data else template_path.parent

    if output_path.exists() and not args.no_backup:
        bak_file = create_backup(output_path)
        print(f"Backup created: {bak_file.name}")

    res = render_template(template_path, output_path, context, base_dir)

    print(f"SUCCESS: Rendered template '{template_path.name}' -> '{output_path.name}'.")
    print(f"  Destination: {output_path}")
    print(f"  Variables provided: {', '.join(res['variables_rendered']) if res['variables_rendered'] else 'None'}")


if __name__ == "__main__":
    main()
