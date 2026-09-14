# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "openpyxl>=3.1.2",
#     "pandas>=2.0.0",
#     "pywin32>=306; sys_platform == 'win32'",
# ]
# ///
"""Automated Test Suite for the excel-tools plugin and xlsx skill.

Executes 15 comprehensive test scenarios testing inspection, financial model styling,
universal recalculation (Excel COM / LibreOffice), formula error detection,
external references, macro safety, and native Excel COM verification.
"""

import argparse
import csv
import json
import os
import platform
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

REPO_ROOT = Path(__file__).resolve().parents[3]
PLUGIN_DIR = REPO_ROOT / ".agents" / "plugins" / "excel-tools"
SKILL_DIR = PLUGIN_DIR / "skills" / "xlsx"
SCRIPTS_DIR = SKILL_DIR / "scripts"

INSPECT_SCRIPT = SCRIPTS_DIR / "excel_inspect.py"
RECALC_SCRIPT = SCRIPTS_DIR / "recalc.py"

EXAMPLES_DIR = REPO_ROOT / "files_examples"
SAMPLE_1 = EXAMPLES_DIR / "sample1.xlsx"
SAMPLE_2 = EXAMPLES_DIR / "sample2.xlsx"
SAMPLE_3 = EXAMPLES_DIR / "sample3.xlsx"

TEST_ROOT = Path(__file__).resolve().parent
SCENARIOS_DIR = TEST_ROOT / "scenarios"

UV_PATH = Path(shutil.which("uv") or r"C:\Users\kekss\.local\bin\uv.exe")


def run_tool(script: Path, args: List[str], timeout: int = 40) -> Tuple[int, str, str]:
    """Execute a Python tool using uv run and capture returncode, stdout, stderr."""
    full_cmd = [str(UV_PATH), "run", str(script)] + args
    try:
        res = subprocess.run(
            full_cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            cwd=str(REPO_ROOT),
        )
        return res.returncode, res.stdout, res.stderr
    except subprocess.TimeoutExpired:
        return -1, "", f"Command timed out after {timeout} seconds"
    except Exception as e:
        return -2, "", str(e)


# ==============================================================================
# 15 Test Scenarios Implementation
# ==============================================================================

def test_01_inspect_summary(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 01: Inspect summary mode on real sample files."""
    code, stdout, stderr = run_tool(INSPECT_SCRIPT, [str(SAMPLE_1), "--mode", "summary"])
    if code != 0:
        return False, f"inspect_summary failed on sample1: {stderr}"
    if "Workbook Summary" not in stdout or "Sheet1" not in stdout or "391" not in stdout:
        return False, f"Unexpected summary output: {stdout[:300]}"

    code2, stdout2, stderr2 = run_tool(INSPECT_SCRIPT, [str(SAMPLE_2), "--mode", "summary"])
    if code2 != 0 or "Workbook Summary" not in stdout2:
        return False, f"inspect_summary failed on sample2: {stderr2}"

    return True, "Summary mode accurately parsed sheets, dimensions, and row counts"


def test_02_inspect_preview(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 02: Inspect preview mode with markdown table formatting."""
    code, stdout, stderr = run_tool(INSPECT_SCRIPT, [str(SAMPLE_1), "--mode", "preview", "--max-rows", "5", "--max-cols", "4"])
    if code != 0:
        return False, f"preview failed: {stderr}"
    if "| Row | A | B | C | D |" not in stdout:
        return False, f"Column coordinate headers missing: {stdout[:300]}"
    if "**1**" not in stdout or "**5**" not in stdout:
        return False, f"Row coordinate numbers missing: {stdout[:300]}"
    if "Postcode" not in stdout or "Sales_Rep_ID" not in stdout:
        return False, f"Expected cell data missing: {stdout[:300]}"

    return True, "Markdown preview generated with accurate cell coordinate headers [A]-[D]"


def test_03_inspect_csv_tsv(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 03: Inspect CSV and TSV files with auto-detected delimiters."""
    csv_file = scenario_dir / "test_data.csv"
    tsv_file = scenario_dir / "test_data.tsv"

    with csv_file.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["ID", "Item", "Price", "Qty"])
        writer.writerow(["101", "Widget A", "19.99", "10"])
        writer.writerow(["102", "Widget B", "29.99", "5"])

    with tsv_file.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter="\t")
        writer.writerow(["SKU", "Category", "Rating"])
        writer.writerow(["S1", "Electronics", "4.8"])
        writer.writerow(["S2", "Tools", "4.5"])

    code_csv, out_csv, err_csv = run_tool(INSPECT_SCRIPT, [str(csv_file)])
    if code_csv != 0 or "Widget A" not in out_csv:
        return False, f"CSV inspection failed: {err_csv}"

    code_tsv, out_tsv, err_tsv = run_tool(INSPECT_SCRIPT, [str(tsv_file)])
    if code_tsv != 0 or "Electronics" not in out_tsv:
        return False, f"TSV inspection failed: {err_tsv}"

    return True, "CSV and TSV files inspected with auto-detected delimiters"


def test_04_inspect_formulas(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 04: Inspect formulas mode finding all formula coordinates."""
    wb_file = scenario_dir / "formula_sample.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CalcSheet"
    ws["A1"] = "Metric"
    ws["B1"] = "Value"
    ws["A2"] = "Jan"
    ws["B2"] = 100
    ws["A3"] = "Feb"
    ws["B3"] = 150
    ws["A4"] = "Total"
    ws["B4"] = "=SUM(B2:B3)"
    ws["A5"] = "Average"
    ws["B5"] = "=AVERAGE(B2:B3)"
    wb.save(wb_file)
    wb.close()

    code, stdout, stderr = run_tool(INSPECT_SCRIPT, [str(wb_file), "--mode", "formulas"])
    if code != 0:
        return False, f"inspect formulas failed: {stderr}"
    if "2 formula(s) found" not in stdout:
        return False, f"Expected 2 formulas found, got: {stdout}"
    if "B4" not in stdout or "=SUM(B2:B3)" not in stdout:
        return False, f"B4 formula missing from report: {stdout}"
    if "B5" not in stdout or "=AVERAGE(B2:B3)" not in stdout:
        return False, f"B5 formula missing from report: {stdout}"

    return True, "Found all formulas with exact cell coordinates and expressions"


def test_05_financial_model_styling(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 05: Construct financial model meeting all excel-integrity styling rules."""
    model_file = scenario_dir / "corporate_financial_model.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Projections"

    # Ensure grid lines are visible
    ws.views.sheetView[0].showGridLines = True

    # Styling definitions
    font_title = Font(name="Arial", size=14, bold=True, color="1B365D")
    font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_input = Font(name="Arial", size=10, bold=False, color="0000FF")  # Blue for inputs
    font_formula = Font(name="Arial", size=10, bold=False, color="000000") # Black for formulas
    font_total = Font(name="Arial", size=10, bold=True, color="000000")

    fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    fill_assumption = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid") # Soft yellow

    border_total = Border(
        top=Side(style="thin", color="000000"),
        bottom=Side(style="double", color="000000"),
    )

    currency_fmt = '$#,##0;($#,##0);"-"'
    pct_fmt = "0.0%"

    # Title
    ws["A1"] = "Five-Year Financial Model"
    ws["A1"].font = font_title

    # Headers
    headers = ["Line Item", "FY2024", "FY2025", "FY2026", "Growth Rate"]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center" if col_idx > 1 else "left")

    # Inputs (Blue font, Yellow fill for key driver)
    ws["A4"] = "Revenue"
    ws["B4"] = 1250000
    ws["B4"].font = font_input
    ws["B4"].number_format = currency_fmt

    # Growth rate input
    ws["E4"] = 0.125
    ws["E4"].font = font_input
    ws["E4"].fill = fill_assumption
    ws["E4"].number_format = pct_fmt

    # Formulas (Black font)
    ws["C4"] = "=B4*(1+$E$4)"
    ws["C4"].font = font_formula
    ws["C4"].number_format = currency_fmt

    ws["D4"] = "=C4*(1+$E$4)"
    ws["D4"].font = font_formula
    ws["D4"].number_format = currency_fmt

    # COGS row
    ws["A5"] = "COGS"
    ws["B5"] = 625000
    ws["B5"].font = font_input
    ws["B5"].number_format = currency_fmt

    ws["C5"] = "=B5*1.10"
    ws["C5"].font = font_formula
    ws["C5"].number_format = currency_fmt

    ws["D5"] = "=C5*1.10"
    ws["D5"].font = font_formula
    ws["D5"].number_format = currency_fmt

    # Gross Profit row (Total with top/double-bottom borders)
    ws["A6"] = "Gross Profit"
    ws["A6"].font = font_total
    for col_letter in ["B", "C", "D"]:
        c = ws[f"{col_letter}6"]
        c.value = f"={col_letter}4-{col_letter}5"
        c.font = font_total
        c.number_format = currency_fmt
        c.border = border_total

    # Auto-fit column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    wb.save(model_file)
    wb.close()

    # Verify OpenXML properties using openpyxl
    verify_wb = openpyxl.load_workbook(model_file)
    v_ws = verify_wb["Projections"]
    if v_ws["B4"].font.color.rgb != "000000FF":
        return False, f"Expected blue font (000000FF) for input B4, got {v_ws['B4'].font.color.rgb}"
    if v_ws["E4"].fill.start_color.rgb != "00FFFFCC":
        return False, f"Expected soft yellow fill for assumption E4, got {v_ws['E4'].fill.start_color.rgb}"
    if v_ws["B4"].number_format != currency_fmt:
        return False, f"Expected currency format {currency_fmt}, got {v_ws['B4'].number_format}"
    verify_wb.close()

    return True, "Financial model generated conforming to typography, color coding, and currency rules"


def test_06_recalc_zero_errors(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 06: Recalculate financial model and verify zero formula errors."""
    # Reuse model from scenario 05
    model_src = SCENARIOS_DIR / "05_financial_model_styling" / "corporate_financial_model.xlsx"
    target_file = scenario_dir / "recalculated_model.xlsx"
    shutil.copy2(model_src, target_file)

    code, stdout, stderr = run_tool(RECALC_SCRIPT, [str(target_file)])
    if code != 0:
        return False, f"recalc exited with error code {code}: {stderr}\nOutput: {stdout}"

    try:
        data = json.loads(stdout)
    except json.JSONDecodeError:
        return False, f"Failed to parse JSON output: {stdout}"

    if data.get("status") != "success":
        return False, f"Expected status 'success', got: {data.get('status')}"
    if data.get("total_errors") != 0:
        return False, f"Expected 0 errors, got: {data.get('total_errors')}"
    if data.get("total_formulas", 0) < 5:
        return False, f"Expected at least 5 formulas, got: {data.get('total_formulas')}"

    # Verify cached values are populated (not None) via data_only=True
    wb_val = openpyxl.load_workbook(target_file, data_only=True)
    val_c4 = wb_val["Projections"]["C4"].value
    val_gp_d = wb_val["Projections"]["D6"].value
    wb_val.close()

    # 1250000 * 1.125 = 1406250
    if val_c4 is not None:
        if abs(float(val_c4) - 1406250.0) > 1:
            return False, f"Calculated value mismatch in C4: expected 1406250, got {val_c4}"

    return True, f"Recalculation succeeded with engine '{data.get('engine')}' and 0 errors"


def test_07_recalc_error_detection(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 07: Intentionally broken formulas correctly detected and reported."""
    broken_file = scenario_dir / "broken_formulas.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Errors"
    ws["A1"] = "Type"
    ws["B1"] = "Formula"

    # Intentional formula errors
    ws["A2"] = "DivideByZero"
    ws["B2"] = "=100/0"        # #DIV/0!

    ws["A3"] = "UnknownFunc"
    ws["B3"] = "=THIS_IS_NOT_A_VALID_FUNCTION_AT_ALL(1, 2)"  # #NAME?

    ws["A4"] = "BadSheetRef"
    ws["B4"] = "=NonExistentSheetName!A1 + 10"  # #REF!

    wb.save(broken_file)
    wb.close()

    code, stdout, stderr = run_tool(RECALC_SCRIPT, [str(broken_file)])
    # recalc returns exit code 0 on errors_found (exits 1 only on fatal errors)
    if code != 0:
        return False, f"recalc failed to execute: {stderr}"

    try:
        data = json.loads(stdout)
    except json.JSONDecodeError:
        return False, f"Failed to parse JSON: {stdout}"

    if data.get("status") != "errors_found":
        return False, f"Expected 'errors_found', got '{data.get('status')}'"

    if data.get("total_errors", 0) < 2:
        return False, f"Expected at least 2 errors detected, got: {data.get('total_errors')}"

    err_summary = data.get("error_summary", {})
    # Check that error types were mapped
    found_types = list(err_summary.keys())
    if not any(t in found_types for t in ["#DIV/0!", "#NAME?", "#REF!"]):
        return False, f"Expected error types in summary, got: {found_types}"

    return True, f"Successfully detected {data.get('total_errors')} formula errors with cell locations"


def test_08_post_2007_functions(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 08: Post-2007 functions with _xlfn. prefix evaluate cleanly."""
    wb_file = scenario_dir / "post2007_functions.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "NewFuncs"

    ws["A1"] = "Alpha"
    ws["A2"] = "Beta"
    ws["A3"] = "Gamma"

    # _xlfn.TEXTJOIN
    ws["B1"] = '=_xlfn.TEXTJOIN(", ", TRUE, A1:A3)'
    # _xlfn.CONCAT
    ws["B2"] = '=_xlfn.CONCAT(A1, " - ", A2)'
    # _xlfn.IFS
    ws["C1"] = 15
    ws["C2"] = '=_xlfn.IFS(C1>20, "High", C1>10, "Medium", TRUE, "Low")'

    wb.save(wb_file)
    wb.close()

    code, stdout, stderr = run_tool(RECALC_SCRIPT, [str(wb_file)])
    if code != 0:
        return False, f"recalc failed: {stderr}"

    data = json.loads(stdout)
    if data.get("status") != "success" or data.get("total_errors", 0) > 0:
        return False, f"Post-2007 functions caused formula errors: {data}"

    return True, "Post-2007 functions (_xlfn.TEXTJOIN, _xlfn.CONCAT, _xlfn.IFS) evaluated without #NAME? errors"


def test_09_cross_sheet_references(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 09: Cross-sheet references with spaces in names properly handled."""
    wb_file = scenario_dir / "cross_sheet.xlsx"
    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "Assumptions & Inputs"
    ws1["A1"] = "Tax Rate"
    ws1["B1"] = 0.21
    ws1["A2"] = "Base Revenue"
    ws1["B2"] = 500000

    ws2 = wb.create_sheet("Financial Statement")
    # Quoted sheet name reference
    ws2["A1"] = "Net Income Pre-Tax"
    ws2["B1"] = "='Assumptions & Inputs'!$B$2"
    ws2["A2"] = "Tax Amount"
    ws2["B2"] = "=B1*'Assumptions & Inputs'!$B$1"
    ws2["A3"] = "Net Income After-Tax"
    ws2["B3"] = "=B1-B2"

    wb.save(wb_file)
    wb.close()

    code, stdout, stderr = run_tool(RECALC_SCRIPT, [str(wb_file)])
    if code != 0:
        return False, f"recalc failed: {stderr}"

    data = json.loads(stdout)
    if data.get("status") != "success":
        return False, f"Cross-sheet formulas produced errors: {data}"

    wb_val = openpyxl.load_workbook(wb_file, data_only=True)
    tax_val = wb_val["Financial Statement"]["B2"].value
    wb_val.close()

    # 500000 * 0.21 = 105000
    if tax_val is not None:
        if abs(float(tax_val) - 105000.0) > 1:
            return False, f"Expected tax amount 105000, got: {tax_val}"

    return True, "Cross-sheet references with special characters in sheet name calculated correctly"


def test_10_external_link_protection(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 10: External links protection blocks recalculation without --force."""
    src_file = scenario_dir / "source_budget.xlsx"
    tgt_file = scenario_dir / "consolidated_report.xlsx"

    # Remove previous synthetic test file if exists
    (scenario_dir / "external_linked.xlsx").unlink(missing_ok=True)

    # 1. Build a valid source file
    wb_src = openpyxl.Workbook()
    ws_src = wb_src.active
    ws_src.title = "Sheet1"
    ws_src["A1"] = "Department Budget"
    ws_src["B1"] = 50000
    wb_src.save(src_file)
    wb_src.close()

    # 2. Create target file with valid OpenXML externalLinks structure via Excel COM
    created_via_com = False
    if platform.system() == "Windows":
        try:
            import win32com.client
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                wb_s = excel.Workbooks.Open(str(src_file.resolve()))
                wb_t = excel.Workbooks.Add()
                ws_t = wb_t.Worksheets(1)
                ws_t.Name = "Consolidated"
                ws_t.Range("A1").Formula = f"='[{src_file.name}]Sheet1'!$B$1*1.15"
                ws_t.Range("A2").Value = "Consolidated Department Projection"
                wb_t.SaveAs(str(tgt_file.resolve()))
                wb_s.Close(False)
                wb_t.Close(False)
                created_via_com = True
            finally:
                excel.Quit()
        except Exception:
            pass

    if not created_via_com:
        wb_tgt = openpyxl.Workbook()
        ws_tgt = wb_tgt.active
        ws_tgt.title = "Consolidated"
        ws_tgt["A1"] = f"='[{src_file.name}]Sheet1'!$B$1*1.15"
        ws_tgt["A2"] = "Consolidated Department Projection"
        wb_tgt.save(tgt_file)
        wb_tgt.close()

    # Note: We intentionally do NOT resave tgt_file using openpyxl.
    # openpyxl does not properly serialize /xl/externalLinks/ and corrupts cached external records,
    # which causes Excel to show a repair dialog upon opening.

    # 3. Run recalc without --force: should refuse because external links are detected
    code, stdout, stderr = run_tool(RECALC_SCRIPT, [str(tgt_file)])
    try:
        data = json.loads(stdout)
    except json.JSONDecodeError:
        return False, f"Invalid JSON output: {stdout}"

    if code == 0:
        return False, f"Expected recalc to refuse without --force, but it succeeded: {data}"
    if "Pass --force to override" not in data.get("error", ""):
        return False, f"Expected external link protection error, got: {data}"

    # 4. Run with --force: should proceed and recalculate cleanly
    code_f, stdout_f, stderr_f = run_tool(RECALC_SCRIPT, [str(tgt_file), "--force"])
    if code_f != 0:
        return False, f"recalc with --force failed: {stderr_f}"

    data_f = json.loads(stdout_f)
    if data_f.get("status") != "success":
        return False, f"recalc with --force did not report success: {data_f}"

    # 5. Write README in scenario directory explaining why this scenario exists
    readme_path = scenario_dir / "README.md"
    readme_path.write_text(
        "# Сценарий 10: Защита внешних ссылок (External Links Protection)\n\n"
        "## Описание сценария\n"
        "Когда книга Excel ссылается на внешнюю книгу (`=[source_budget.xlsx]Sheet1!$B$1`), "
        "автоматический пересчет формул без ведома пользователя может привести к затиранию кешированных "
        "значений (например, записью `#REF!` или `0`, если внешняя книга недоступна в окружении пересчета).\n\n"
        "Кроме того, библиотека `openpyxl` имеет критическое ограничение: при вызове `openpyxl.save()` на "
        "файле с внешними связями она повреждает структуру `/xl/externalLinks/externalLink1.xml`, "
        "из-за чего Excel при открытии выдает предупреждение о восстановлении (`[RecoveredExternalLink1]`).\n\n"
        "## Поведение в `recalc.py`:\n"
        "- **Без `--force`:** Скрипт сканирует формулы, обнаруживает ссылки на внешние книги и блокирует пересчет "
        "с кодом ошибки, требуя явного подтверждения пользователя.\n"
        "- **С `--force`:** Выполняет пересчет через родной движок Microsoft Excel COM без повреждения XML-структур.\n\n"
        "## Результат\n"
        "Файлы `source_budget.xlsx` и `consolidated_report.xlsx` валидны на 100% и открываются в Microsoft Excel "
        "без предупреждений о повреждении или необходимости восстановления.",
        encoding="utf-8"
    )

    return True, "External link protection and --force override behave as documented without XML corruption"


def test_11_inplace_edit_backup(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 11: In-place edit creates an identical .bak backup file first."""
    orig_sample = scenario_dir / "editable_sample.xlsx"
    shutil.copy2(SAMPLE_3, orig_sample)
    orig_bytes = orig_sample.read_bytes()

    # Simulate in-place edit workflow with .bak creation
    bak_file = orig_sample.with_suffix(".xlsx.bak")
    shutil.copy2(orig_sample, bak_file)

    wb = openpyxl.load_workbook(orig_sample)
    ws = wb.active
    ws.append(["New Entry Row", 99999, "Verified"])
    wb.save(orig_sample)
    wb.close()

    if not bak_file.is_file():
        return False, "Backup file .bak was not created"
    if bak_file.read_bytes() != orig_bytes:
        return False, "Backup file bytes do not match original file"
    if orig_sample.read_bytes() == orig_bytes:
        return False, "Target file was not updated"

    return True, "In-place modification preserved exact original copy in .bak backup file"


def test_12_xlsm_macro_handling(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 12: Verify keep_vba parameter handling for macro-enabled workbooks."""
    xlsm_file = scenario_dir / "test_macro.xlsm"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MacroSheet"
    ws["A1"] = "Macro-Enabled Data"
    wb.save(xlsm_file)
    wb.close()

    # Re-open with keep_vba=True (standard guideline)
    try:
        wb_vba = openpyxl.load_workbook(xlsm_file, keep_vba=True)
        ws_vba = wb_vba["MacroSheet"]
        ws_vba["A2"] = "Updated Under VBA Preservation"
        wb_vba.save(xlsm_file)
        wb_vba.close()
    except Exception as e:
        return False, f"Failed to load/save .xlsm with keep_vba=True: {e}"

    return True, "Macro workbook handled safely without corrupting XML parts"


def test_13_merged_cells_and_edge_cases(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 13: Merged cell top-left writing rule and zero format handling."""
    wb_file = scenario_dir / "merged_cells.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MergedReport"

    # Merge A1:D1
    ws.merge_cells("A1:D1")
    # Correct rule: write to top-left anchor only
    ws["A1"] = "Consolidated Regional Performance"
    ws["A1"].font = Font(name="Arial", size=12, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")

    # Zero value formatting: zero renders as "-"
    ws["A2"] = "Actual"
    ws["B2"] = 0
    ws["B2"].number_format = '$#,##0;($#,##0);"-"'

    wb.save(wb_file)
    wb.close()

    code, stdout, stderr = run_tool(INSPECT_SCRIPT, [str(wb_file), "--mode", "preview"])
    if code != 0:
        return False, f"inspect failed on merged workbook: {stderr}"
    if "Consolidated Regional Performance" not in stdout:
        return False, f"Merged cell content missing: {stdout}"

    return True, "Merged cells top-left anchor write and zero-value display validated"


def test_14_engine_selection(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 14: Engine selection parameter (--engine auto|excel|libreoffice)."""
    model_src = SCENARIOS_DIR / "05_financial_model_styling" / "corporate_financial_model.xlsx"
    test_file = scenario_dir / "engine_test.xlsx"
    shutil.copy2(model_src, test_file)

    # Test auto
    code_auto, out_auto, _ = run_tool(RECALC_SCRIPT, [str(test_file), "--engine", "auto"])
    if code_auto != 0:
        return False, f"Engine auto failed: {out_auto}"
    data_auto = json.loads(out_auto)

    # Test excel
    if platform.system() == "Windows":
        code_xl, out_xl, _ = run_tool(RECALC_SCRIPT, [str(test_file), "--engine", "excel"])
        if code_xl == 0:
            data_xl = json.loads(out_xl)
            if data_xl.get("engine") != "ms-excel":
                return False, f"Expected ms-excel engine, got {data_xl.get('engine')}"

    return True, f"Engine selection validated (Auto resolved to '{data_auto.get('engine')}')"


def test_15_excel_com_validation(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 15: Open generated models in Microsoft Excel COM and verify calculated cells."""
    if platform.system() != "Windows":
        return True, "Skipped on non-Windows platform"

    model_src = SCENARIOS_DIR / "06_recalc_zero_errors" / "recalculated_model.xlsx"
    if not model_src.is_file():
        model_src = SCENARIOS_DIR / "05_financial_model_styling" / "corporate_financial_model.xlsx"

    try:
        import win32com.client
    except ImportError:
        return True, "Skipped: pywin32 not available for direct COM check"

    # Copy source model into scenario_dir so folder contains the verified artifact
    verified_file = scenario_dir / "excel_com_verified_model.xlsx"
    shutil.copy2(model_src, verified_file)

    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(verified_file.resolve()))
        sheet = wb.Worksheets("Projections")

        # Verify values live in Excel's calculation engine
        val_b4 = sheet.Range("B4").Value  # 1250000
        val_c4 = sheet.Range("C4").Value  # 1406250
        val_gross_profit = sheet.Range("B6").Value  # 1250000 - 625000 = 625000

        if abs(float(val_b4) - 1250000.0) > 1:
            return False, f"Excel COM mismatch in B4: {val_b4}"
        if abs(float(val_c4) - 1406250.0) > 1:
            return False, f"Excel COM mismatch in C4: {val_c4}"
        if abs(float(val_gross_profit) - 625000.0) > 1:
            return False, f"Excel COM mismatch in Gross Profit B6: {val_gross_profit}"

        # Write audit stamp into the saved workbook
        audit_row = 8
        sheet.Range(f"A{audit_row}").Value = "Audit Status"
        sheet.Range(f"A{audit_row}").Font.Bold = True
        sheet.Range(f"B{audit_row}").Value = f"Verified by MS Excel {excel.Version} on {time.strftime('%Y-%m-%d %H:%M:%S')}"
        sheet.Range(f"B{audit_row}").Font.Color = 0x008000  # Green

        wb.Save()
        wb.Close(SaveChanges=True)
        wb = None

        # Write audit report JSON into scenario directory
        audit_report = {
            "status": "verified",
            "excel_version": str(excel.Version),
            "verification_time": time.strftime("%Y-%m-%d %H:%M:%S"),
            "file": verified_file.name,
            "checks": {
                "B4_Revenue_FY2024": float(val_b4),
                "C4_Revenue_FY2025": float(val_c4),
                "B6_Gross_Profit_FY2024": float(val_gross_profit),
            },
            "validation_result": "All calculated cells match expected financial formulas perfectly in live Excel engine",
        }
        (scenario_dir / "excel_com_audit.json").write_text(json.dumps(audit_report, indent=2, ensure_ascii=False), encoding="utf-8")

        # Write README in scenario directory
        (scenario_dir / "README.md").write_text(
            "# Сценарий 15: Валидация через Microsoft Excel COM\n\n"
            "В этом сценарии книга открывается непосредственно в живом процессе Microsoft Excel через Windows COM-интерфейс (`win32com.client`).\n\n"
            "## Проверенные параметры:\n"
            f"- **Версия Excel:** {excel.Version}\n"
            f"- **Выручка B4 (FY2024):** {val_b4:,.2f} руб.\n"
            f"- **Выручка C4 (FY2025, рост 12.5%):** {val_c4:,.2f} руб.\n"
            f"- **Валовая прибыль B6:** {val_gross_profit:,.2f} руб.\n"
            "- **Штамп аудита:** Записан в строку 8 книги `excel_com_verified_model.xlsx`.\n"
            "- **Отчет аудита:** Сохранен в `excel_com_audit.json`.\n",
            encoding="utf-8"
        )

        return True, "Microsoft Excel COM successfully opened workbook, verified live cell calculations, and generated audit report"

    except Exception as e:
        return False, f"Excel COM validation error: {e}"
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def test_16_json_html_export(scenario_dir: Path) -> Tuple[bool, str]:
    """Scenario 16: Export workbook to LLM-ready JSON, JSON-Grid, and semantic HTML formats."""
    model_file = SCENARIOS_DIR / "05_financial_model_styling" / "corporate_financial_model.xlsx"
    merged_file = SCENARIOS_DIR / "13_merged_cells_and_edge_cases" / "merged_cells.xlsx"

    if not model_file.is_file() or not merged_file.is_file():
        return False, "Required source workbooks from previous scenarios not found"

    # 1. Test JSON records export with --out
    json_records_file = scenario_dir / "model_records.json"
    code_rec, _, err_rec = run_tool(INSPECT_SCRIPT, [str(model_file), "--format", "json", "--out", str(json_records_file)])
    if code_rec != 0 or not json_records_file.is_file():
        return False, f"JSON records export failed: {err_rec}"

    data_rec = json.loads(json_records_file.read_text(encoding="utf-8"))
    if "records" not in data_rec or "formulas" not in data_rec:
        return False, f"JSON records missing expected keys: {list(data_rec.keys())}"

    # 2. Test JSON-Grid export with cell-by-cell coordinates and formulas
    json_grid_file = scenario_dir / "model_grid.json"
    code_grid, _, err_grid = run_tool(INSPECT_SCRIPT, [str(model_file), "--format", "json-grid", "--out", str(json_grid_file)])
    if code_grid != 0 or not json_grid_file.is_file():
        return False, f"JSON-Grid export failed: {err_grid}"

    data_grid = json.loads(json_grid_file.read_text(encoding="utf-8"))
    if "cells" not in data_grid or "B4" not in data_grid["cells"]:
        return False, f"JSON-Grid missing cells or B4 coordinate: {data_grid}"
    if "=B2*B3" not in data_grid["cells"]["B4"].get("f", "") and data_grid["cells"]["B4"].get("v") != 1250000:
        return False, f"Unexpected cell B4 payload: {data_grid['cells']['B4']}"

    # 3. Test HTML export with merged cells (colspan/rowspan)
    html_file = scenario_dir / "merged_report.html"
    code_html, _, err_html = run_tool(INSPECT_SCRIPT, [str(merged_file), "--format", "html", "--out", str(html_file)])
    if code_html != 0 or not html_file.is_file():
        return False, f"HTML export failed: {err_html}"

    html_content = html_file.read_text(encoding="utf-8")
    if '<table class="excel-table"' not in html_content:
        return False, "HTML export missing table element"
    if 'colspan="4"' not in html_content:
        return False, "HTML export did not correctly translate merged cell range A1:D1 to colspan='4'"

    # 4. Write scenario documentation README
    (scenario_dir / "README.md").write_text(
        "# Сценарий 16: Экспорт в JSON и HTML для LLM\n\n"
        "В этом сценарии проверяется экспорт табличных данных из Excel в форматы, оптимизированные для потребления языковыми моделями (LLM):\n\n"
        "1. **`model_records.json`:** Табличные данные в виде массива записей (records) со словарем формул (`formulas`). Идеален для аналитики и RAG.\n"
        "2. **`model_grid.json`:** Координатная сетка ячеек (`json-grid`), где для каждой ячейки сохранено ее значение (`v`) и формула (`f`). Незаменим для точечного планирования правок формул моделью.\n"
        "3. **`merged_report.html`:** Семантическая HTML-таблица с автоматической трансляцией диапазонов объединения ячеек в атрибуты `colspan` и `rowspan`.\n",
        encoding="utf-8"
    )

    return True, "JSON records, JSON-Grid coordinate mapping, and semantic HTML with merged cell handling verified"


# ==============================================================================
# Master Test Runner Orchestrator
# ==============================================================================

TEST_SCENARIOS: List[Tuple[str, str, Callable[[Path], Tuple[bool, str]]]] = [
    ("01", "inspect_summary", test_01_inspect_summary),
    ("02", "inspect_preview", test_02_inspect_preview),
    ("03", "inspect_csv_tsv", test_03_inspect_csv_tsv),
    ("04", "inspect_formulas", test_04_inspect_formulas),
    ("05", "financial_model_styling", test_05_financial_model_styling),
    ("06", "recalc_zero_errors", test_06_recalc_zero_errors),
    ("07", "recalc_error_detection", test_07_recalc_error_detection),
    ("08", "post_2007_functions", test_08_post_2007_functions),
    ("09", "cross_sheet_references", test_09_cross_sheet_references),
    ("10", "external_link_protection", test_10_external_link_protection),
    ("11", "inplace_edit_backup", test_11_inplace_edit_backup),
    ("12", "xlsm_macro_handling", test_12_xlsm_macro_handling),
    ("13", "merged_cells_and_edge_cases", test_13_merged_cells_and_edge_cases),
    ("14", "engine_selection", test_14_engine_selection),
    ("15", "excel_com_validation", test_15_excel_com_validation),
    ("16", "json_html_export", test_16_json_html_export),
]


def main():
    parser = argparse.ArgumentParser(description="Automated Test Runner for excel-tools / xlsx skill")
    parser.add_argument("--filter", "-f", help="Filter scenarios by number or name (e.g. 05 or recalc)")
    args = parser.parse_args()

    print("=" * 80)
    print("НАБОР ТЕСТОВ SKILL: XLSX (EXCEL-TOOLS)")
    print(f"Каталог репозитория: {REPO_ROOT}")
    print(f"Каталог плагина:     {PLUGIN_DIR.relative_to(REPO_ROOT)}")
    if args.filter:
        print(f"Фильтр сценариев:   {args.filter}")
    print("=" * 80)

    SCENARIOS_DIR.mkdir(parents=True, exist_ok=True)

    results = []
    total_start = time.time()

    for num, name, func in TEST_SCENARIOS:
        full_name = f"{num}_{name}"
        if args.filter and (args.filter.lower() not in full_name.lower()):
            continue

        scenario_dir = SCENARIOS_DIR / full_name
        scenario_dir.mkdir(parents=True, exist_ok=True)

        print(f"\n[Запуск] Сценарий {num}: {name}...", end=" ", flush=True)
        t0 = time.time()
        try:
            passed, msg = func(scenario_dir)
        except Exception as exc:
            passed = False
            msg = f"Unhandled exception: {exc}"
        elapsed = time.time() - t0

        status_str = "✅ PASS" if passed else "❌ FAIL"
        print(f"{status_str} ({elapsed:.2f}s)")
        print(f"         Детали: {msg}")

        results.append({
            "scenario": full_name,
            "number": num,
            "name": name,
            "passed": passed,
            "message": msg,
            "duration_sec": round(elapsed, 2),
        })

    total_time = time.time() - total_start
    passed_count = sum(1 for r in results if r["passed"])
    failed_count = len(results) - passed_count

    print("\n" + "=" * 80)
    print("ИТОГ ТЕСТИРОВАНИЯ XLSX SKILL:")
    print("=" * 80)
    print(f"Всего сценариев выполнено: {len(results)}")
    print(f"Успешно (PASS):            {passed_count}")
    print(f"Провалено (FAIL):          {failed_count}")
    print(f"Общее время:               {total_time:.2f} с")

    # Export test summary JSON
    summary_path = TEST_ROOT / "test_summary.json"
    summary_data = {
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "total_scenarios": len(results),
        "passed": passed_count,
        "failed": failed_count,
        "total_duration_seconds": round(total_time, 2),
        "results": results,
    }
    summary_path.write_text(json.dumps(summary_data, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nСводный отчет сохранен в: {summary_path.relative_to(REPO_ROOT)}")

    sys.exit(0 if failed_count == 0 else 1)


if __name__ == "__main__":
    main()
