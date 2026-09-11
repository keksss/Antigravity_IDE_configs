# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pywin32>=306; sys_platform == 'win32'",
#     "docx2pdf>=0.1.8; sys_platform == 'win32' or sys_platform == 'darwin'",
# ]
# ///
"""Convert Microsoft Word (.docx) documents to PDF using native Word COM or LibreOffice headless."""

import argparse
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import List, Optional, Tuple

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def is_valid_docx(file_path: Path) -> bool:
    """Check if the file exists and is a valid DOCX archive."""
    if not file_path.is_file():
        return False
    if file_path.name.startswith("~$"):
        return False  # Word lock temporary file
    try:
        with zipfile.ZipFile(file_path, "r") as z:
            return "word/document.xml" in z.namelist()
    except (zipfile.BadZipFile, OSError):
        return False


def find_libreoffice_binary() -> Optional[Path]:
    """Locate the LibreOffice / soffice binary across platforms."""
    # 1. Check custom environment variable
    custom_path = os.environ.get("LIBREOFFICE_PATH")
    if custom_path and Path(custom_path).is_file():
        return Path(custom_path)

    # 2. Check standard system PATH
    for candidate in ["soffice", "libreoffice", "soffice.exe", "libreoffice.exe"]:
        found = shutil.which(candidate)
        if found:
            return Path(found)

    # 3. Check known default installation paths by platform
    system = platform.system()
    search_paths = []
    if system == "Windows":
        prog_files = os.environ.get("ProgramFiles", r"C:\Program Files")
        prog_files_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
        search_paths = [
            Path(prog_files) / "LibreOffice" / "program" / "soffice.exe",
            Path(prog_files_x86) / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "LibreOffice" / "program" / "soffice.exe",
        ]
    elif system == "Darwin":  # macOS
        search_paths = [
            Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
            Path("~/Applications/LibreOffice.app/Contents/MacOS/soffice").expanduser(),
        ]
    elif system == "Linux":
        search_paths = [
            Path("/usr/bin/soffice"),
            Path("/usr/bin/libreoffice"),
            Path("/usr/local/bin/soffice"),
            Path("/usr/local/bin/libreoffice"),
            Path("/snap/bin/libreoffice"),
        ]

    for p in search_paths:
        if p.is_file():
            return p

    return None


def convert_with_word(input_file: Path, output_file: Path, timeout_sec: int = 20) -> bool:
    """Convert DOCX to PDF using Word in an isolated subprocess with timeout and process cleanup."""
    system = platform.system()

    if system == "Darwin":
        try:
            from docx2pdf import convert
            convert(str(input_file.resolve()), str(output_file.resolve()))
            return output_file.exists() and output_file.stat().st_size > 0
        except Exception as e:
            print(f"[Word Error] {e}", file=sys.stderr)
            return False

    if system != "Windows":
        return False

    # Run isolated Python script with win32com to prevent hanging parent process on modal dialogs
    worker_code = (
        "import sys\n"
        "from pathlib import Path\n"
        "try:\n"
        "    import win32com.client\n"
        "    w = win32com.client.DispatchEx('Word.Application')\n"
        "    w.Visible = False\n"
        "    w.DisplayAlerts = 0\n"
        "    doc = w.Documents.Open(str(Path(sys.argv[1]).resolve()), ReadOnly=True, ConfirmConversions=False, NoEncodingDialog=True)\n"
        "    doc.SaveAs(str(Path(sys.argv[2]).resolve()), FileFormat=17)\n"
        "    doc.Close(False)\n"
        "    w.Quit()\n"
        "    sys.exit(0)\n"
        "except Exception as e:\n"
        "    sys.stderr.write(str(e))\n"
        "    sys.exit(1)\n"
    )

    try:
        res = subprocess.run(
            [sys.executable, "-c", worker_code, str(input_file.resolve()), str(output_file.resolve())],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=timeout_sec,
        )
        if res.returncode == 0 and output_file.exists() and output_file.stat().st_size > 0:
            return True
        else:
            err = res.stderr.strip() or res.stdout.strip()
            if err:
                print(f"[Word COM Error] {err}", file=sys.stderr)
            return False
    except subprocess.TimeoutExpired:
        print(f"[Word COM Warning] Word did not respond within {timeout_sec}s (likely waiting on activation or modal dialog).", file=sys.stderr)
        try:
            subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
        except Exception:
            pass
        return False
    except Exception as e:
        print(f"[Word Error] {e}", file=sys.stderr)
        return False


def convert_with_libreoffice(soffice_path: Path, input_file: Path, output_file: Path) -> bool:
    """Convert DOCX to PDF using LibreOffice headless with an isolated profile."""
    output_dir = output_file.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="libreoffice_pdf_") as temp_profile_dir:
        profile_uri = Path(temp_profile_dir).resolve().as_uri()
        cmd = [
            str(soffice_path),
            f"-env:UserInstallation={profile_uri}",
            "--headless",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            str(output_dir.resolve()),
            str(input_file.resolve()),
        ]

        try:
            res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=120)
            expected_name = input_file.stem + ".pdf"
            default_output = output_dir / expected_name

            if default_output.exists():
                if default_output.resolve() != output_file.resolve():
                    if output_file.exists():
                        output_file.unlink()
                    default_output.rename(output_file)
                return output_file.exists() and output_file.stat().st_size > 0
            else:
                print(f"[LibreOffice Error] Output file not found. Output:\n{res.stdout}\n{res.stderr}", file=sys.stderr)
                return False
        except Exception as e:
            print(f"[LibreOffice Error] Execution failed: {e}", file=sys.stderr)
            return False


def convert_single_file(
    input_file: Path,
    output_file: Optional[Path] = None,
    engine: str = "auto",
    force: bool = False,
    soffice_path: Optional[Path] = None,
) -> Tuple[bool, str]:
    """Convert one DOCX file to PDF using the requested engine."""
    input_file = input_file.resolve()
    if not is_valid_docx(input_file):
        return False, f"Invalid or locked DOCX file: {input_file}"

    if output_file is None:
        output_file = input_file.with_suffix(".pdf")
    else:
        output_file = output_file.resolve()

    if output_file.exists() and not force:
        return False, f"Output file already exists (use --force to overwrite): {output_file}"

    output_file.parent.mkdir(parents=True, exist_ok=True)

    chosen_engine = engine.lower()

    if chosen_engine == "word":
        success = convert_with_word(input_file, output_file)
        return (success, f"Converted with Microsoft Word -> {output_file}" if success else "Word conversion failed.")

    elif chosen_engine == "libreoffice":
        if not soffice_path:
            return False, "LibreOffice executable (soffice) not found on the system."
        success = convert_with_libreoffice(soffice_path, input_file, output_file)
        return (success, f"Converted with LibreOffice -> {output_file}" if success else "LibreOffice conversion failed.")

    elif chosen_engine == "auto":
        # 1. Try Word first if on Windows/macOS
        if platform.system() in ("Windows", "Darwin"):
            success = convert_with_word(input_file, output_file)
            if success:
                return True, f"Converted with Microsoft Word -> {output_file}"
            print("[Warning] Word conversion unavailable, attempting LibreOffice...", file=sys.stderr)

        # 2. Try LibreOffice
        if soffice_path:
            success = convert_with_libreoffice(soffice_path, input_file, output_file)
            if success:
                return True, f"Converted with LibreOffice -> {output_file}"

        return False, (
            "No conversion engine succeeded.\n"
            "Requirements:\n"
            "  - Microsoft Word installed and activated (Windows/macOS), OR\n"
            "  - LibreOffice installed with 'soffice' accessible in PATH or default directories."
        )

    else:
        return False, f"Unknown engine: {engine}"


def collect_docx_files(paths: List[Path]) -> List[Path]:
    """Collect all valid .docx files from provided files or folders."""
    files: List[Path] = []
    for p in paths:
        if p.is_dir():
            for item in p.rglob("*.docx"):
                if not item.name.startswith("~$"):
                    files.append(item)
        elif p.is_file():
            if not p.name.startswith("~$"):
                files.append(p)
        else:
            print(f"[Warning] Path not found or inaccessible: {p}", file=sys.stderr)
    return sorted(list(set(files)))


def main():
    parser = argparse.ArgumentParser(
        description="Convert Microsoft Word (.docx) files to PDF using native Word or LibreOffice."
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        type=Path,
        help="One or more .docx files, or directories containing .docx files.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Target .pdf file path (for single file) or destination directory (for batch mode).",
    )
    parser.add_argument(
        "-e",
        "--engine",
        choices=["auto", "word", "libreoffice"],
        default="auto",
        help="Conversion backend to use (default: auto).",
    )
    parser.add_argument(
        "-f",
        "--force",
        action="store_true",
        help="Overwrite existing output PDF files without prompting.",
    )

    args = parser.parse_args()

    # Pre-detect LibreOffice if needed
    soffice_path = find_libreoffice_binary()

    files = collect_docx_files(args.inputs)
    if not files:
        print("[Error] No valid .docx files found to convert.", file=sys.stderr)
        sys.exit(1)

    # Check if single file with specific output file path
    is_single_file_mode = len(files) == 1 and args.output and not args.output.is_dir() and args.output.suffix.lower() == ".pdf"

    success_count = 0
    failure_count = 0

    if is_single_file_mode:
        input_file = files[0]
        output_file = args.output
        print(f"Converting '{input_file.name}' -> '{output_file.name}' (engine: {args.engine})...")
        ok, msg = convert_single_file(
            input_file=input_file,
            output_file=output_file,
            engine=args.engine,
            force=args.force,
            soffice_path=soffice_path,
        )
        if ok:
            print(f"[SUCCESS] {msg}")
            success_count += 1
        else:
            print(f"[ERROR] {msg}", file=sys.stderr)
            failure_count += 1
    else:
        # Batch or directory mode
        out_dir = args.output if (args.output and (args.output.is_dir() or args.output.suffix == "")) else None
        if out_dir:
            out_dir.mkdir(parents=True, exist_ok=True)

        print(f"Found {len(files)} document(s) to convert (engine: {args.engine}).")

        for idx, docx_file in enumerate(files, 1):
            if out_dir:
                target_pdf = out_dir / (docx_file.stem + ".pdf")
            else:
                target_pdf = docx_file.with_suffix(".pdf")

            print(f"[{idx}/{len(files)}] Converting: {docx_file.name} ...")
            ok, msg = convert_single_file(
                input_file=docx_file,
                output_file=target_pdf,
                engine=args.engine,
                force=args.force,
                soffice_path=soffice_path,
            )
            if ok:
                print(f"  -> [SUCCESS] {target_pdf.name}")
                success_count += 1
            else:
                print(f"  -> [FAILED] {msg}", file=sys.stderr)
                failure_count += 1

    print(f"\nDone: {success_count} succeeded, {failure_count} failed.")
    if failure_count > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
