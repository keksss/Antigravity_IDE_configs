# excel-tools Plugin for Google Antigravity

A modular, portable Antigravity plugin for creating, inspecting, editing, formatting, and recalculating Microsoft Excel (`.xlsx`, `.xlsm`) and tabular (`.csv`, `.tsv`) spreadsheets with zero formula errors, OpenXML integrity, and professional financial model styling.

---

## Features

- **Isolated Dependency Management with `uv`:** All scripts use PEP 723 metadata (`# /// script ...`) to run with zero manual package installation and zero pollution of the global Python environment.
- **Universal Recalculator (`recalc.py`):** Multi-engine recalculation using native Microsoft Excel COM automation (on Windows) or LibreOffice headless (cross-platform), ensuring cached formula values are evaluated and zero runtime errors (`#VALUE!`, `#REF!`, `#NAME?`) exist before delivery.
- **Built-in Inspector (`excel_inspect.py`):** Inspect sheet dimensions, formula counts, column headers, and render tabular Markdown grids with cell coordinate headers without needing external tools.
- **Strict Domain Rules (`excel-integrity.md`):** Enforces `.bak` backups before in-place modifications, formula preservation, dynamic formulas over hardcoded numbers, and executive financial styling.

---

## Directory Structure

```text
excel-tools/
├── plugin.json                       # Plugin manifest metadata
├── README.md                         # This documentation
├── rules/
│   └── excel-integrity.md            # Spreadsheet safety & financial model rules
└── skills/
    └── xlsx/                         # Core spreadsheet creation & editing skill
        ├── SKILL.md                  # Skill instructions and best practices
        └── scripts/
            ├── recalc.py             # Universal recalculator (Excel COM / LibreOffice)
            └── excel_inspect.py      # Sheet inspector & Markdown table previewer
```

---

## Skills in this Plugin

| Skill | Directory | Description |
| :--- | :--- | :--- |
| **`xlsx`** | [`skills/xlsx/`](./skills/xlsx/SKILL.md) | Universal skill for reading, creating, modifying, and recalculating spreadsheets (`.xlsx`, `.xlsm`, `.csv`, `.tsv`), financial models, and tabular data. |

---

## Domain Rules

- [`rules/excel-integrity.md`](./rules/excel-integrity.md): Defines safety standards, `.bak` backup protocols, zero formula error enforcement, openpyxl gotchas, and financial modeling conventions (color palettes, number formatting, and formula compatibility).

---

## Quick Usage Guide

### 1. Inspect a Spreadsheet
Inspect sheets, row/col counts, and formulas:
```bash
# High-level summary of all sheets
uv run skills/xlsx/scripts/excel_inspect.py model.xlsx --mode summary

# Markdown table preview of first 20 rows with cell coordinates [A], [B]
uv run skills/xlsx/scripts/excel_inspect.py model.xlsx --mode preview --max-rows 20

# List all formula expressions
uv run skills/xlsx/scripts/excel_inspect.py model.xlsx --mode formulas
```

### 2. Recalculate Formulas & Validate Errors
Run the universal recalculation engine to compute cached values and check for errors:
```bash
# Auto-detect best engine (MS Excel COM on Windows, LibreOffice headless fallback)
uv run skills/xlsx/scripts/recalc.py model.xlsx

# Force MS Excel engine
uv run skills/xlsx/scripts/recalc.py model.xlsx --engine excel

# Force LibreOffice engine
uv run skills/xlsx/scripts/recalc.py model.xlsx --engine libreoffice
```

Sample output:
```json
{
  "status": "success",
  "engine": "ms-excel",
  "total_formulas": 14,
  "total_errors": 0,
  "error_summary": {}
}
```

---

## Environment Setup

All scripts require `uv` (recommended):
```bash
uv run skills/xlsx/scripts/<script_name>.py [args]
```

Dependencies are declared inline in each script header via PEP 723:
- `openpyxl>=3.1.2`
- `pandas>=2.0.0`
- `pywin32>=306; sys_platform == 'win32'`
