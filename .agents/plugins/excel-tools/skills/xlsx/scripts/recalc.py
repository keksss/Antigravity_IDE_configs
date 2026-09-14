# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "openpyxl>=3.1.2",
#     "pywin32>=306; sys_platform == 'win32'",
# ]
# ///
"""Universal Excel Formula Recalculation & Error Verification Engine.

Recalculates all formulas in an Excel (.xlsx / .xlsm) workbook and inspects for runtime errors.
Supports native Microsoft Excel (via COM automation on Windows) and LibreOffice Calc (headless cross-platform),
with automatic engine selection and graceful diagnostics.
"""

import argparse
import contextlib
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

MAX_LOCATIONS = 100

EXCEL_ERRORS = [
    "#VALUE!",
    "#DIV/0!",
    "#REF!",
    "#NAME?",
    "#NULL!",
    "#NUM!",
    "#N/A",
]

EXTERNAL_REF_RE = re.compile(r"""(?<![\w"\[])'?\[(?:(?:\d+)|(?:[^\]]+\.xlsx?))\][^!"\[]*'?\!""", re.IGNORECASE)

STARBASIC_RECALCULATE_MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>
"""


def get_file_stamp(file_path: Path) -> Tuple[int, int]:
    """Return mtime_ns and file size for change detection."""
    stat = file_path.stat()
    return stat.st_mtime_ns, stat.st_size


def find_libreoffice_binary() -> Optional[Path]:
    """Locate the LibreOffice / soffice executable across platforms."""
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
        local_appdata = os.environ.get("LOCALAPPDATA", "")
        search_paths = [
            Path(prog_files) / "LibreOffice" / "program" / "soffice.exe",
            Path(prog_files_x86) / "LibreOffice" / "program" / "soffice.exe",
            Path(local_appdata) / "Programs" / "LibreOffice" / "program" / "soffice.exe",
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


def is_excel_com_available() -> bool:
    """Check if Microsoft Excel COM automation is available on Windows."""
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def recalculate_with_excel_com(file_path: Path, timeout_seconds: int = 30) -> Tuple[bool, Optional[str]]:
    """Recalculate formulas in an Excel workbook using native MS Excel COM automation."""
    if platform.system() != "Windows":
        return False, "Microsoft Excel COM is only supported on Windows"

    try:
        import win32com.client
    except ImportError:
        return False, "pywin32 library is not available in the environment"

    excel = None
    workbook = None
    abs_path_str = str(file_path.resolve())

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        workbook = excel.Workbooks.Open(abs_path_str)
        # Full recalculation across all open worksheets
        excel.CalculateFull()
        workbook.Save()
        workbook.Close(SaveChanges=True)
        workbook = None
        return True, None

    except Exception as exc:
        return False, f"Excel COM error: {exc}"

    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def setup_libreoffice_profile(profile_dir: Path, soffice_bin: Path, timeout: int = 15) -> Tuple[Optional[str], Optional[str]]:
    """Initialize a clean temporary LibreOffice user profile with the StarBasic recalculation macro."""
    url = profile_dir.as_uri()
    try:
        subprocess.run(
            [
                str(soffice_bin),
                "--headless",
                "--terminate_after_init",
                f"-env:UserInstallation={url}",
            ],
            capture_output=True,
            timeout=timeout,
            check=False,
        )
    except FileNotFoundError:
        return None, f"LibreOffice binary not found at: {soffice_bin}"
    except subprocess.TimeoutExpired:
        return None, "LibreOffice timed out during profile initialization"

    macro_dir = profile_dir / "user" / "basic" / "Standard"
    if not macro_dir.exists():
        # Attempt to create directory structure if soffice init didn't create standard module
        macro_dir.mkdir(parents=True, exist_ok=True)

    macro_file = macro_dir / "Module1.xba"
    try:
        macro_file.write_text(STARBASIC_RECALCULATE_MACRO, encoding="utf-8")
    except OSError as exc:
        return None, f"Could not write StarBasic macro: {exc}"

    return url, None


def recalculate_with_libreoffice(file_path: Path, soffice_bin: Path, timeout_seconds: int = 30) -> Tuple[bool, Optional[str]]:
    """Recalculate formulas in an Excel workbook using LibreOffice headless."""
    abs_path_str = str(file_path.resolve())
    before_stamp = get_file_stamp(file_path)

    with tempfile.TemporaryDirectory(prefix="recalc-lo-profile-", ignore_cleanup_errors=True) as temp_dir:
        profile_dir = Path(temp_dir)
        profile_url, err = setup_libreoffice_profile(profile_dir, soffice_bin, timeout=min(15, timeout_seconds))
        if err:
            return False, err

        cmd = [
            str(soffice_bin),
            "--headless",
            "--norestore",
            f"-env:UserInstallation={profile_url}",
            "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application",
            abs_path_str,
        ]

        try:
            res = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=timeout_seconds,
            )
            if res.returncode != 0:
                detail = (res.stderr or "").strip() or f"Process exited with code {res.returncode}"
                return False, f"LibreOffice failed: {detail}"

        except subprocess.TimeoutExpired:
            return False, f"LibreOffice recalculation timed out after {timeout_seconds} seconds"
        except FileNotFoundError:
            return False, "LibreOffice executable not found"

    after_stamp = get_file_stamp(file_path)
    if after_stamp == before_stamp:
        # Note: If no formulas changed or needed calculation, timestamps might match,
        # but usually store() updates mtime.
        pass

    return True, None


def check_external_links_at_risk(file_path: Path) -> List[str]:
    """Detect external workbook references (e.g. ='[1]Sheet'!$A$1) that risk corruption."""
    try:
        with zipfile.ZipFile(file_path, "r") as archive:
            names = archive.namelist()
    except (zipfile.BadZipFile, OSError):
        return []

    if not any(name.startswith("xl/externalLinks/") for name in names):
        return []

    from openpyxl import load_workbook

    at_risk: List[str] = []
    with contextlib.ExitStack() as stack:
        try:
            wb_formulas = load_workbook(file_path, data_only=False)
            stack.callback(wb_formulas.close)
            wb_values = load_workbook(file_path, data_only=True)
            stack.callback(wb_values.close)
        except Exception:
            return []

        for sheet_name in wb_formulas.sheetnames:
            ws_formulas = wb_formulas[sheet_name]
            if not hasattr(ws_formulas, "iter_rows"):
                continue
            ws_values = wb_values[sheet_name]
            for row in ws_formulas.iter_rows():
                for cell in row:
                    val = cell.value
                    if isinstance(val, str) and val.startswith("="):
                        if EXTERNAL_REF_RE.search(val):
                            at_risk.append(f"{sheet_name}!{cell.coordinate}")

    return at_risk


def inspect_formula_errors(file_path: Path) -> Tuple[int, int, Dict[str, Dict]]:
    """Inspect evaluated workbook for standard Excel formula errors using openpyxl."""
    from openpyxl import load_workbook

    error_details: Dict[str, List[str]] = {err: [] for err in EXCEL_ERRORS}
    total_errors = 0

    # 1. Inspect values for formula errors
    wb_val = load_workbook(file_path, data_only=True)
    try:
        for sheet_name in wb_val.sheetnames:
            ws = wb_val[sheet_name]
            if not hasattr(ws, "iter_rows"):
                continue
            for row in ws.iter_rows():
                for cell in row:
                    val = cell.value
                    if val is not None and isinstance(val, str):
                        for err in EXCEL_ERRORS:
                            if err in val:
                                error_details[err].append(f"{sheet_name}!{cell.coordinate}")
                                total_errors += 1
                                break
    finally:
        wb_val.close()

    error_summary = {}
    for err_type, locations in error_details.items():
        if locations:
            entry = {
                "count": len(locations),
                "locations": locations[:MAX_LOCATIONS],
            }
            if len(locations) > MAX_LOCATIONS:
                entry["locations_truncated"] = len(locations) - MAX_LOCATIONS
            error_summary[err_type] = entry

    # 2. Count total formulas defined in workbook
    wb_form = load_workbook(file_path, data_only=False)
    formula_count = 0
    try:
        for sheet_name in wb_form.sheetnames:
            ws = wb_form[sheet_name]
            if not hasattr(ws, "iter_rows"):
                continue
            for row in ws.iter_rows():
                for cell in row:
                    val = cell.value
                    if isinstance(val, str) and val.startswith("="):
                        formula_count += 1
    finally:
        wb_form.close()

    return formula_count, total_errors, error_summary


def recalculate(file_path: Path, timeout: int = 30, force: bool = False, preferred_engine: str = "auto") -> Dict:
    """Universal recalculation coordinator."""
    if not file_path.is_file():
        return {"error": f"File does not exist: {file_path}"}

    if not os.access(file_path, os.W_OK):
        return {"error": f"File is not writable: {file_path}"}

    # Check external links protection
    if not force:
        try:
            at_risk = check_external_links_at_risk(file_path)
            if at_risk:
                return {
                    "error": (
                        "Refusing to recalculate: workbook contains external file references that would lose "
                        f"cached values ({len(at_risk)} cell(s)). Pass --force to override."
                    ),
                    "external_link_cells": at_risk[:MAX_LOCATIONS],
                }
        except Exception as exc:
            return {"error": f"Failed to check external links: {exc}"}

    engine_used = "none"
    recalc_success = False
    recalc_err = None

    excel_available = is_excel_com_available()
    lo_binary = find_libreoffice_binary()

    # Engine selection logic
    if preferred_engine == "excel":
        if excel_available:
            recalc_success, recalc_err = recalculate_with_excel_com(file_path, timeout)
            if recalc_success:
                engine_used = "ms-excel"
        else:
            recalc_err = "MS Excel COM requested but not available"

    elif preferred_engine == "libreoffice":
        if lo_binary:
            recalc_success, recalc_err = recalculate_with_libreoffice(file_path, lo_binary, timeout)
            if recalc_success:
                engine_used = "libreoffice"
        else:
            recalc_err = "LibreOffice requested but not found"

    else:  # "auto" mode: prefer native Excel on Windows, fallback to LibreOffice
        if platform.system() == "Windows" and excel_available:
            recalc_success, recalc_err = recalculate_with_excel_com(file_path, timeout)
            if recalc_success:
                engine_used = "ms-excel"
            elif lo_binary:
                # Fallback to LibreOffice if Excel COM failed
                recalc_success, recalc_err = recalculate_with_libreoffice(file_path, lo_binary, timeout)
                if recalc_success:
                    engine_used = "libreoffice"
        elif lo_binary:
            recalc_success, recalc_err = recalculate_with_libreoffice(file_path, lo_binary, timeout)
            if recalc_success:
                engine_used = "libreoffice"

    # Now inspect formula status via openpyxl
    try:
        formula_count, total_errors, error_summary = inspect_formula_errors(file_path)
    except Exception as exc:
        return {"error": f"Error inspecting formulas in workbook: {exc}"}

    result = {
        "status": "success" if total_errors == 0 else "errors_found",
        "engine": engine_used,
        "total_formulas": formula_count,
        "total_errors": total_errors,
        "error_summary": error_summary,
    }

    if engine_used == "none":
        result["engine_warning"] = (
            "Neither Microsoft Excel nor LibreOffice was available for background calculation. "
            "Formulas are preserved in the workbook and will calculate automatically when opened in Excel/Calc."
        )
        if recalc_err:
            result["engine_detail"] = recalc_err

    return result


def main():
    parser = argparse.ArgumentParser(
        description="Universal Excel formula recalculation and error verification script."
    )
    parser.add_argument("file", type=Path, help="Path to the .xlsx or .xlsm file to recalculate")
    parser.add_argument(
        "--timeout",
        type=int,
        default=30,
        help="Timeout in seconds for recalculation engine (default: 30)",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Force recalculation even if external links might lose cached values",
    )
    parser.add_argument(
        "--engine",
        choices=["auto", "excel", "libreoffice"],
        default="auto",
        help="Preferred calculation engine (default: auto)",
    )

    args = parser.parse_args()

    result = recalculate(
        file_path=args.file,
        timeout=args.timeout,
        force=args.force,
        preferred_engine=args.engine,
    )

    print(json.dumps(result, indent=2))
    sys.exit(1 if "error" in result else 0)


if __name__ == "__main__":
    main()
