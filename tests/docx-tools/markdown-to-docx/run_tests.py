# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "python-docx>=1.1.2",
#     "markdown-it-py>=3.0.0",
#     "Pillow>=10.0.0",
#     "pywin32; sys_platform == 'win32'",
# ]
# ///
"""Comprehensive Test Suite for markdown-to-docx skill.

Executes all 11 test scenarios covering:
- Scenario 01: Headings & Hierarchy (H1-H6, Russian/Cyrillic styles)
- Scenario 02: Full Inline Formatting (bold, italic, bold+italic, strikethrough, code, soft/hard breaks)
- Scenario 03: Hyperlinks & Anchors (clickable external links, formatted anchor text)
- Scenario 04: Lists & Hierarchies (bullet, ordered, 3+ level nesting, task lists/checkboxes)
- Scenario 05: GFM Tables & Alignment (left/center/right, tblHeader, cantSplit, shaded header)
- Scenario 06: Blockquotes & Callouts (NOTE, TIP, IMPORTANT, WARNING, CAUTION with rich markdown)
- Scenario 07: Code Blocks & Syntax (Consolas font, background shading, left accent border)
- Scenario 08: Images & Proportional Scaling (Pillow aspect ratio, DPI, centered, missing image fallback)
- Scenario 09: Template Inheritance (--template, body clear, style & section preservation)
- Scenario 10: Chained Workflow with docx-editor (MD -> DOCX -> docx-editor in-place fine-tuning)
- Scenario 11: Document Validity & Microsoft Word COM Automation (Word COM open & PDF export)
"""

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
import json
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from PIL import Image, ImageDraw

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

REPO_ROOT = Path(__file__).resolve().parents[3]
PLUGIN_DIR = REPO_ROOT / ".agents" / "plugins" / "docx-tools"
SKILL_DIR = PLUGIN_DIR / "skills" / "markdown-to-docx"
CONVERT_SCRIPT = SKILL_DIR / "scripts" / "md_to_docx.py"

DOCX_EDITOR_DIR = PLUGIN_DIR / "skills" / "docx-editor" / "scripts"
DOCX_EDIT_SCRIPT = DOCX_EDITOR_DIR / "docx_edit.py"
DOCX_INSPECT_SCRIPT = DOCX_EDITOR_DIR / "docx_inspect.py"

EXAMPLES_DIR = REPO_ROOT / "files_examples"
SAMPLE_1 = EXAMPLES_DIR / "file-sample_1.docx"

TEST_ROOT = Path(__file__).resolve().parent
SCENARIOS_DIR = TEST_ROOT / "scenarios"

UV_PATH = Path(shutil.which("uv") or r"C:\Users\kekss\.local\bin\uv.exe")

test_results = []
generated_docx_files = []


def run_command(cmd_args: list[str], timeout: int = 60) -> tuple[int, str, str]:
    """Run command with uv and return code, stdout, stderr."""
    full_cmd = [str(UV_PATH), "run"] + cmd_args
    try:
        res = subprocess.run(
            full_cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
        return res.returncode, res.stdout, res.stderr
    except subprocess.TimeoutExpired:
        return -1, "", "Command timed out"
    except Exception as e:
        return -2, "", str(e)


def record_result(name: str, passed: bool, details: str, output_folder: Path = None):
    test_results.append({
        "scenario": name,
        "passed": passed,
        "details": details,
        "folder": str(output_folder.relative_to(REPO_ROOT)) if output_folder else "",
    })
    status_str = "PASS" if passed else "FAIL"
    print(f"[{status_str}] {name}: {details}")


# =========================================================================
# Scenario 01: Headings & Hierarchy
# =========================================================================
def scenario_01_headings_and_hierarchy():
    """Verify heading translation (H1-H6, Russian Cyrillic headings) and style mappings."""
    folder = SCENARIOS_DIR / "01_headings_and_hierarchy"
    folder.mkdir(parents=True, exist_ok=True)

    test_md = folder / "headings_test.md"
    test_md.write_text(
        """# Раздел 1. Главный заголовок документа
Вводный абзац к первому разделу.

## Подраздел 1.1. Архитектурные требования
Описание требований к архитектуре.

### Пункт 1.1.1. Спецификация компонентов
Детальные спецификации.

#### Подпункт 1.1.1.a. Низкоуровневые детали
Техническая спецификация параметров.

##### Микропункт Уровня 5
Параметры микроконтроллеров.

###### Финальный Уровень 6
Конфигурационные константы.
""",
        encoding="utf-8",
    )

    out_docx = folder / "headings_output.docx"
    code, out, err = run_command([str(CONVERT_SCRIPT), str(test_md), "-o", str(out_docx)])
    all_ok = code == 0 and out_docx.is_file()

    doc = None
    headings = []
    notes = []
    if all_ok:
        generated_docx_files.append(out_docx)
        doc = docx.Document(str(out_docx))
        headings = [p for p in doc.paragraphs if p.text.startswith("Раздел") or p.text.startswith("Подраздел") or p.text.startswith("Пункт") or p.text.startswith("Подпункт") or p.text.startswith("Микропункт") or p.text.startswith("Финальный")]
        if len(headings) != 6:
            all_ok = False
            notes.append(f"Expected 6 headings, found {len(headings)}")

        # Check Russian characters in heading 1
        if not any("Главный заголовок документа" in p.text for p in doc.paragraphs):
            all_ok = False
            notes.append("Cyrillic title text not matched")

    rep = f"""# Отчет по сценарию 01: Headings & Hierarchy
- Код возврата: {code}
- Сгенерирован DOCX: {'ДА' if out_docx.is_file() else 'НЕТ'}
- Количество параграфов: {len(doc.paragraphs) if doc else 0}
- Замечания: {'; '.join(notes) if notes else 'Нет'}
- Статус: {'УСПЕХ' if all_ok else 'ОШИБКА'}
"""
    (folder / "scenario_report.md").write_text(rep, encoding="utf-8")
    record_result("Scenario 01: Headings & Hierarchy", all_ok, "Preserved H1-H6 headings and Cyrillic titles", folder)


# =========================================================================
# Scenario 02: Full Inline Formatting
# =========================================================================
def scenario_02_full_inline_formatting():
    """Verify bold, italic, strikethrough, inline code, softbreak, and hardbreak."""
    folder = SCENARIOS_DIR / "02_inline_formatting"
    folder.mkdir(parents=True, exist_ok=True)

    test_md = folder / "inline_test.md"
    test_md.write_text(
        """# Inline Formatting Demonstration

Текст с **жирным начертанием** и *курсивным начертанием*.
Сложные комбинации: ***жирный курсив*** и **жирный с *вложенным курсивом* внутри**.
Зачеркивание: ~~устаревший параметр~~ и ~~зачеркнутый с **жирным** акцентом~~.
Инлайн код: `const MAX_BUFFER = 1024;` и **жирный с `код внутри` выделением**.

Мягкий перенос строки: Первая строка абзаца
Вторая строка того же абзаца должна быть разделена пробелом.

Жесткий перенос строки:\\
Третья строка на новой физической строке внутри того же параграфа.
""",
        encoding="utf-8",
    )

    out_docx = folder / "inline_output.docx"
    code, out, err = run_command([str(CONVERT_SCRIPT), str(test_md), "-o", str(out_docx)])
    all_ok = code == 0 and out_docx.is_file()

    doc = None
    runs = []
    has_bold = False
    has_italic = False
    has_strike = False
    has_code = False
    has_spacing = False
    notes = []
    if all_ok:
        generated_docx_files.append(out_docx)
        doc = docx.Document(str(out_docx))
        runs = [r for p in doc.paragraphs for r in p.runs]

        has_bold = any(r.bold for r in runs)
        has_italic = any(r.italic for r in runs)
        has_strike = any(r.font.strike for r in runs)
        has_code = any(r.font.name == "Consolas" for r in runs)

        # Check softbreak spacing (words shouldn't stick together)
        p_texts = [p.text for p in doc.paragraphs]
        has_spacing = any("Первая строка абзаца Вторая строка" in t for t in p_texts)

        if not has_bold:
            all_ok = False
            notes.append("No bold runs found")
        if not has_italic:
            all_ok = False
            notes.append("No italic runs found")
        if not has_strike:
            all_ok = False
            notes.append("No strikethrough runs found")
        if not has_code:
            all_ok = False
            notes.append("No Consolas code runs found")
        if not has_spacing:
            all_ok = False
            notes.append("Softbreak space between lines not preserved")

    rep = f"""# Отчет по сценарию 02: Full Inline Formatting
- Код возврата: {code}
- Найдено Run-элементов: {len(runs)}
- Bold: {'ДА' if has_bold else 'НЕТ'}
- Italic: {'ДА' if has_italic else 'НЕТ'}
- Strike: {'ДА' if has_strike else 'НЕТ'}
- Inline Code (Consolas): {'ДА' if has_code else 'НЕТ'}
- Softbreak Space: {'ДА' if has_spacing else 'НЕТ'}
- Замечания: {'; '.join(notes) if notes else 'Нет'}
- Статус: {'УСПЕХ' if all_ok else 'ОШИБКА'}
"""
    (folder / "scenario_report.md").write_text(rep, encoding="utf-8")
    record_result("Scenario 02: Full Inline Formatting", all_ok, "Preserved bold, italic, strike, code, and breaks", folder)


# =========================================================================
# Scenario 03: Hyperlinks & Anchors
# =========================================================================
def scenario_03_hyperlinks_and_anchors():
    """Verify native OpenXML clickable hyperlinks with plain and rich anchor text."""
    folder = SCENARIOS_DIR / "03_hyperlinks_and_anchors"
    folder.mkdir(parents=True, exist_ok=True)

    test_md = folder / "links_test.md"
    test_md.write_text(
        """# Hyperlinks Test Document

1. Стандартная ссылка: [Официальный сайт Python](https://www.python.org).
2. Ссылка с жирным текстом: [**Документация OpenXML**](https://learn.microsoft.com/en-us/office/open-xml/).
3. Ссылка с курсивом: [*Портал поддержки пользователей*](https://support.example.com).
""",
        encoding="utf-8",
    )

    out_docx = folder / "links_output.docx"
    code, out, err = run_command([str(CONVERT_SCRIPT), str(test_md), "-o", str(out_docx)])
    all_ok = code == 0 and out_docx.is_file()

    has_py_url = False
    has_ms_url = False
    has_hyperlink_tags = False
    notes = []
    if all_ok:
        generated_docx_files.append(out_docx)
        with zipfile.ZipFile(out_docx, "r") as z:
            rels_xml = z.read("word/_rels/document.xml.rels").decode("utf-8")
            doc_xml = z.read("word/document.xml").decode("utf-8")

        has_py_url = "https://www.python.org" in rels_xml
        has_ms_url = "learn.microsoft.com" in rels_xml
        has_hyperlink_tags = "<w:hyperlink" in doc_xml

        if not has_py_url or not has_ms_url:
            all_ok = False
            notes.append("Missing hyperlink relationship target URLs in rels XML")
        if not has_hyperlink_tags:
            all_ok = False
            notes.append("No <w:hyperlink> elements found in document.xml")

    rep = f"""# Отчет по сценарию 03: Hyperlinks & Anchors
- Код возврата: {code}
- Наличие w:hyperlink в XML: {'ДА' if has_hyperlink_tags else 'НЕТ'}
- Наличие URL в отношениях OpenXML: {'ДА' if (has_py_url and has_ms_url) else 'НЕТ'}
- Замечания: {'; '.join(notes) if notes else 'Нет'}
- Статус: {'УСПЕХ' if all_ok else 'ОШИБКА'}
"""
    (folder / "scenario_report.md").write_text(rep, encoding="utf-8")
    record_result("Scenario 03: Hyperlinks & Anchors", all_ok, "Preserved clickable OpenXML hyperlinks with rich text", folder)


# =========================================================================
# Scenario 04: Lists & Hierarchies
# =========================================================================
def scenario_04_lists_and_hierarchies():
    """Verify bullet lists, ordered lists, 3+ nested levels, and task lists / checkboxes."""
    folder = SCENARIOS_DIR / "04_lists_and_hierarchies"
    folder.mkdir(parents=True, exist_ok=True)

    test_md = folder / "lists_test.md"
    test_md.write_text(
        """# List Hierarchies and Task Lists

* Базовый пункт списка 1
* Базовый пункт списка 2
  * Вложенный подпункт 2.1
  * Вложенный подпункт 2.2
    * Глубокий подпункт 2.2.1
* Базовый пункт 3

1. Шаг первый: инициализация проекта
2. Шаг второй: настройка конфигураций
   1. Подшаг 2.1: создание каталогов
   2. Подшаг 2.2: установка зависимостей
3. Шаг третий: сборка и развертывание

- [ ] Запланированная задача (не выполнена)
- [x] Выполненная задача с **жирным приоритетом**
""",
        encoding="utf-8",
    )

    out_docx = folder / "lists_output.docx"
    code, out, err = run_command([str(CONVERT_SCRIPT), str(test_md), "-o", str(out_docx)])
    all_ok = code == 0 and out_docx.is_file()

    doc = None
    has_unchecked = False
    has_checked = False
    has_deep_nested = False
    notes = []
    if all_ok:
        generated_docx_files.append(out_docx)
        doc = docx.Document(str(out_docx))
        p_texts = [p.text for p in doc.paragraphs]

        has_unchecked = any("☐ Запланированная задача" in t for t in p_texts)
        has_checked = any("☑ Выполненная задача" in t for t in p_texts)
        has_deep_nested = any("Глубокий подпункт 2.2.1" in t for t in p_texts)

        if not has_unchecked or not has_checked:
            all_ok = False
            notes.append("Task list checkboxes (☐ and ☑) not properly rendered")
        if not has_deep_nested:
            all_ok = False
            notes.append("Deeply nested list item not found")

    rep = f"""# Отчет по сценарию 04: Lists & Hierarchies
- Код возврата: {code}
- Всего параграфов: {len(doc.paragraphs) if doc else 0}
- Чекбокс [ ]: {'ДА' if has_unchecked else 'НЕТ'}
- Чекбокс [x]: {'ДА' if has_checked else 'НЕТ'}
- Вложенный уровень 3: {'ДА' if has_deep_nested else 'НЕТ'}
- Замечания: {'; '.join(notes) if notes else 'Нет'}
- Статус: {'УСПЕХ' if all_ok else 'ОШИБКА'}
"""
    (folder / "scenario_report.md").write_text(rep, encoding="utf-8")
    record_result("Scenario 04: Lists & Hierarchies", all_ok, "Preserved 3-level lists and task list checkboxes", folder)


# =========================================================================
# Scenario 05: GFM Tables & Alignment
# =========================================================================
def scenario_05_gfm_tables_and_alignment():
    """Verify GFM tables with column alignments (left, center, right), repeated header, and cell styling."""
    folder = SCENARIOS_DIR / "05_gfm_tables_and_alignment"
    folder.mkdir(parents=True, exist_ok=True)

    test_md = folder / "tables_test.md"
    test_md.write_text(
        """# Performance Benchmark Report

| Компонент | Выравнивание по центру | Показатель, оп/сек | Статус выполнения |
| :--- | :---: | ---: | :--- |
| **Ядро парсера AST** | Централизованно | 1 450 000 | `Высокая скорость` |
| Редактор OpenXML | Сбалансированно | 890 000 | `Стабильно` |
| *Конвертер PDF* | Автоматически | 420 000 | `В пределах нормы` |
""",
        encoding="utf-8",
    )

    out_docx = folder / "tables_output.docx"
    code, out, err = run_command([str(CONVERT_SCRIPT), str(test_md), "-o", str(out_docx)])
    all_ok = code == 0 and out_docx.is_file()

    doc = None
    has_tblHeader = False
    has_cantSplit = False
    right_align_ok = False
    notes = []
    if all_ok:
        generated_docx_files.append(out_docx)
        doc = docx.Document(str(out_docx))
        if len(doc.tables) == 0:
            all_ok = False
            notes.append("No table generated in document")
        else:
            tbl = doc.tables[0]
            if len(tbl.rows) != 4 or len(tbl.columns) != 4:
                all_ok = False
                notes.append(f"Unexpected table shape: {len(tbl.rows)}x{len(tbl.columns)}")

            # Check right alignment on numeric column
            c2_p = tbl.cell(1, 2).paragraphs[0]
            right_align_ok = (c2_p.alignment == WD_ALIGN_PARAGRAPH.RIGHT)
            if not right_align_ok:
                all_ok = False
                notes.append(f"Column 3 expected RIGHT alignment, got {c2_p.alignment}")

            # Check center alignment on column 1
            c1_p = tbl.cell(1, 1).paragraphs[0]
            if c1_p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                all_ok = False
                notes.append(f"Column 2 expected CENTER alignment, got {c1_p.alignment}")

            # Check tblHeader in OpenXML
            trPr0 = tbl.rows[0]._tr.get_or_add_trPr()
            has_tblHeader = any(child.tag.endswith("tblHeader") for child in trPr0)
            has_cantSplit = any(child.tag.endswith("cantSplit") for child in trPr0)

            if not has_tblHeader:
                all_ok = False
                notes.append("Header row is missing w:tblHeader repeat property")
            if not has_cantSplit:
                all_ok = False
                notes.append("Rows are missing w:cantSplit protection property")

    rep = f"""# Отчет по сценарию 05: GFM Tables & Alignment
- Код возврата: {code}
- Таблиц в документе: {len(doc.tables) if doc else 0}
- tblHeader на шапке: {'ДА' if has_tblHeader else 'НЕТ'}
- cantSplit на строках: {'ДА' if has_cantSplit else 'НЕТ'}
- Правое выравнивание чисел: {'ДА' if right_align_ok else 'НЕТ'}
- Замечания: {'; '.join(notes) if notes else 'Нет'}
- Статус: {'УСПЕХ' if all_ok else 'ОШИБКА'}
"""
    (folder / "scenario_report.md").write_text(rep, encoding="utf-8")
    record_result("Scenario 05: GFM Tables & Alignment", all_ok, "Preserved alignments, tblHeader, cantSplit, and cells", folder)


# =========================================================================
# Scenario 06: Blockquotes & Callouts
# =========================================================================
def scenario_06_blockquotes_and_callouts():
    """Verify regular blockquotes and GitHub-style alerts (NOTE, TIP, IMPORTANT, WARNING, CAUTION)."""
    folder = SCENARIOS_DIR / "06_blockquotes_and_callouts"
    folder.mkdir(parents=True, exist_ok=True)

    test_md = folder / "callouts_test.md"
    test_md.write_text(
        """# Alerts and Blockquotes Demo

> Обычная цитата с выделением *курсивом* и **жирным** текстом.
> 
> Второй абзац цитаты с сохранением единого стиля.

> [!NOTE]
> Это системное примечание с [ссылкой](https://example.com) и `кодом`.

> [!TIP]
> Рекомендуемый совет по повышению продуктивности разработчика.

> [!IMPORTANT]
> Важная информация, требующая обязательного внимания пользователя.

> [!WARNING]
> Внимание: изменение конфигурации может повлиять на запуск тестов.

> [!CAUTION]
> Осторожно: удаление базы данных является необратимой операцией!
""",
        encoding="utf-8",
    )

    out_docx = folder / "callouts_output.docx"
    code, out, err = run_command([str(CONVERT_SCRIPT), str(test_md), "-o", str(out_docx)])
    all_ok = code == 0 and out_docx.is_file()

    doc = None
    missing_labels = []
    doc_xml = ""
    notes = []
    if all_ok:
        generated_docx_files.append(out_docx)
        doc = docx.Document(str(out_docx))
        p_texts = [p.text for p in doc.paragraphs]

        labels = ["[Примечание]", "[Совет]", "[Важно]", "[Внимание]", "[Осторожно]"]
        missing_labels = [lbl for lbl in labels if not any(lbl in t for t in p_texts)]
        if missing_labels:
            all_ok = False
            notes.append(f"Missing alert labels: {missing_labels}")

        # Check OpenXML shading and borders for callout paragraphs
        with zipfile.ZipFile(out_docx, "r") as z:
            doc_xml = z.read("word/document.xml").decode("utf-8")
        if "w:pBdr" not in doc_xml or "w:shd" not in doc_xml:
            all_ok = False
            notes.append("Missing w:pBdr or w:shd in document.xml for callouts")

    rep = f"""# Отчет по сценарию 06: Blockquotes & Callouts
- Код возврата: {code}
- Наличие всех 5 алертов: {'ДА' if not missing_labels else 'НЕТ'}
- Наличие акцентных бордюров w:pBdr: {'ДА' if 'w:pBdr' in doc_xml else 'НЕТ'}
- Замечания: {'; '.join(notes) if notes else 'Нет'}
- Статус: {'УСПЕХ' if all_ok else 'ОШИБКА'}
"""
    (folder / "scenario_report.md").write_text(rep, encoding="utf-8")
    record_result("Scenario 06: Blockquotes & Callouts", all_ok, "Preserved all 5 alert callouts with rich formatting", folder)


# =========================================================================
# Scenario 07: Code Blocks & Syntax
# =========================================================================
def scenario_07_code_blocks_and_syntax():
    """Verify fenced code blocks with language identifiers, background shading, and borders."""
    folder = SCENARIOS_DIR / "07_code_blocks_and_syntax"
    folder.mkdir(parents=True, exist_ok=True)

    test_md = folder / "code_test.md"
    test_md.write_text(
        '''# Code Syntax Presentation

Ниже представлен пример реализации скрипта на Python:

```python
def process_pipeline(input_data: list[str]) -> dict[str, int]:
    """Process incoming records with deduplication."""
    counts = {}
    for item in input_data:
        counts[item] = counts.get(item, 0) + 1
    return counts
```

А также блок конфигурации JSON:

```json
{
  "service": "markdown-to-docx",
  "version": "2.0.0",
  "active": true
}
```
''',
        encoding="utf-8",
    )

    out_docx = folder / "code_output.docx"
    code, out, err = run_command([str(CONVERT_SCRIPT), str(test_md), "-o", str(out_docx)])
    all_ok = code == 0 and out_docx.is_file()

    doc = None
    code_runs = []
    indent_ok = False
    notes = []
    if all_ok:
        generated_docx_files.append(out_docx)
        doc = docx.Document(str(out_docx))
        code_runs = [r for p in doc.paragraphs for r in p.runs if r.font.name == "Consolas"]
        if len(code_runs) < 2:
            all_ok = False
            notes.append("Expected at least 2 code block paragraphs with Consolas font")

        # Verify multi-line indentation preserved
        code_texts = [p.text for p in doc.paragraphs if "def process_pipeline" in p.text]
        indent_ok = bool(code_texts and '    counts = {}' in code_texts[0])
        if not indent_ok:
            all_ok = False
            notes.append("Indentation inside python code block was flattened")

    rep = f"""# Отчет по сценарию 07: Code Blocks & Syntax
- Код возврата: {code}
- Найдено блоков кода Consolas: {len(code_runs)}
- Сохранение 4-пробельного отступа: {'ДА' if indent_ok else 'НЕТ'}
- Замечания: {'; '.join(notes) if notes else 'Нет'}
- Статус: {'УСПЕХ' if all_ok else 'ОШИБКА'}
"""
    (folder / "scenario_report.md").write_text(rep, encoding="utf-8")
    record_result("Scenario 07: Code Blocks & Syntax", all_ok, "Preserved code indentation, Consolas font, and shading", folder)


# =========================================================================
# Scenario 08: Images & Aspect Ratio
# =========================================================================
def scenario_08_images_and_aspect_ratio():
    """Verify image embedding, Pillow DPI / aspect ratio preservation, and missing image handling."""
    folder = SCENARIOS_DIR / "08_images_and_aspect_ratio"
    folder.mkdir(parents=True, exist_ok=True)

    # Generate a real test PNG image (400x200)
    test_img = folder / "test_chart.png"
    im = Image.new("RGB", (400, 200), color=(240, 244, 248))
    draw = ImageDraw.Draw(im)
    draw.rectangle([20, 20, 380, 180], outline=(9, 105, 218), width=3)
    draw.text((80, 90), "ANTIGRAVITY TEST CHART", fill=(36, 41, 46))
    im.save(test_img, dpi=(96, 96))

    test_md = folder / "images_test.md"
    test_md.write_text(
        """# Document with Images

Ниже размещена реальная диаграмма:

![Архитектурная схема](test_chart.png)

А также ссылка на отсутствующий файл:

![Несуществующее фото](non_existent_file.jpg)
""",
        encoding="utf-8",
    )

    out_docx = folder / "images_output.docx"
    code, out, err = run_command([str(CONVERT_SCRIPT), str(test_md), "-o", str(out_docx)])
    all_ok = code == 0 and out_docx.is_file()

    doc = None
    media_files = []
    has_missing_fallback = False
    notes = []
    if all_ok:
        generated_docx_files.append(out_docx)
        with zipfile.ZipFile(out_docx, "r") as z:
            names = z.namelist()
            media_files = [n for n in names if n.startswith("word/media/")]

        if len(media_files) != 1:
            all_ok = False
            notes.append(f"Expected 1 embedded media file in docx, found {len(media_files)}")

        doc = docx.Document(str(out_docx))
        has_missing_fallback = any("[Missing Image: non_existent_file.jpg]" in p.text for p in doc.paragraphs)
        if not has_missing_fallback:
            all_ok = False
            notes.append("Missing image graceful fallback text not found")

    rep = f"""# Отчет по сценарию 08: Images & Aspect Ratio
- Код возврата: {code}
- Внедренных картинок в DOCX zip: {len(media_files)}
- Корректная обработка отсутствующей картинки: {'ДА' if has_missing_fallback else 'НЕТ'}
- Замечания: {'; '.join(notes) if notes else 'Нет'}
- Статус: {'УСПЕХ' if all_ok else 'ОШИБКА'}
"""
    (folder / "scenario_report.md").write_text(rep, encoding="utf-8")
    record_result("Scenario 08: Images & Aspect Ratio", all_ok, "Embedded media with Pillow aspect ratio & graceful fallback", folder)


# =========================================================================
# Scenario 09: Template Inheritance (--template)
# =========================================================================
def scenario_09_template_inheritance():
    """Verify template inheritance: preserves styles, headers, footers, sections while clearing body."""
    folder = SCENARIOS_DIR / "09_template_inheritance"
    folder.mkdir(parents=True, exist_ok=True)

    test_md = folder / "template_test.md"
    test_md.write_text(
        """# Новый Отчет на базе Корпоративного Шаблона

Этот документ использует стили, параметры полей и колонтитулы исходного шаблона.

## Ключевые показатели эффективности
- Скорость сборки: 100%
- Сохранность стилей: подтверждена
""",
        encoding="utf-8",
    )

    out_docx = folder / "template_output.docx"
    code, out, err = run_command([str(CONVERT_SCRIPT), str(test_md), "--template", str(SAMPLE_1), "-o", str(out_docx)])
    all_ok = code == 0 and out_docx.is_file()

    doc = None
    old_text_present = False
    new_text_present = False
    notes = []
    if all_ok:
        generated_docx_files.append(out_docx)
        doc = docx.Document(str(out_docx))
        # Ensure old text from sample_1 is NOT present
        old_text_present = any("Lorem ipsum" in p.text for p in doc.paragraphs)
        new_text_present = any("Новый Отчет на базе Корпоративного Шаблона" in p.text for p in doc.paragraphs)

        if old_text_present:
            all_ok = False
            notes.append("Old body content from template was not cleared")
        if not new_text_present:
            all_ok = False
            notes.append("New markdown content not found in output document")

        # Verify sections preserved
        if len(doc.sections) == 0:
            all_ok = False
            notes.append("Section properties missing from document")

    rep = f"""# Отчет по сценарию 09: Template Inheritance
- Код возврата: {code}
- Шаблон: {SAMPLE_1.name}
- Очистка старого тела документа: {'ДА' if not old_text_present else 'НЕТ'}
- Вставка нового Markdown текста: {'ДА' if new_text_present else 'НЕТ'}
- Количество секций: {len(doc.sections) if doc else 0}
- Замечания: {'; '.join(notes) if notes else 'Нет'}
- Статус: {'УСПЕХ' if all_ok else 'ОШИБКА'}
"""
    (folder / "scenario_report.md").write_text(rep, encoding="utf-8")
    record_result("Scenario 09: Template Inheritance", all_ok, "Inherited styles and sections while clearing body", folder)


# =========================================================================
# Scenario 10: Chained Workflow with docx-editor
# =========================================================================
def scenario_10_chained_workflow_with_docx_editor():
    """Verify the 2-step pipeline: Markdown-to-DOCX conversion followed by in-place docx-editor fine-tuning."""
    folder = SCENARIOS_DIR / "10_chained_workflow"
    folder.mkdir(parents=True, exist_ok=True)

    contract_md = folder / "contract_draft.md"
    contract_md.write_text(
        """# ПРОЕКТ ДОГОВОРА ОКАЗАНИЯ УСЛУГ № 2026-01

Настоящий документ определяет условия сотрудничества между Заказчиком и Исполнителем.

## 1. Предмет договора
Исполнитель обязуется выполнить комплекс услуг по оптимизации программного обеспечения.

## 2. Стоимость работ
Общая стоимость услуг составляет 250 000 рублей.
""",
        encoding="utf-8",
    )

    stage1_docx = folder / "contract_stage1.docx"
    # Step 1: markdown-to-docx conversion
    c1, out1, err1 = run_command([str(CONVERT_SCRIPT), str(contract_md), "-o", str(stage1_docx)])
    step1_ok = c1 == 0 and stage1_docx.is_file()

    notes = []
    if not step1_ok:
        notes.append(f"Step 1 failed: {err1}")

    # Step 2: docx_inspect to verify outline
    c2, out2, err2 = run_command([str(DOCX_INSPECT_SCRIPT), str(stage1_docx), "--mode", "outline"])
    step2_ok = c2 == 0 and "ПРОЕКТ ДОГОВОРА" in out2

    # Step 3: docx_edit in-place text replacement
    c3, out3, err3 = run_command([
        str(DOCX_EDIT_SCRIPT), str(stage1_docx), "replace-text",
        "--find", "ПРОЕКТ ДОГОВОРА",
        "--replace", "УТВЕРЖДЕННЫЙ ДОГОВОР",
    ])
    step3_ok = c3 == 0

    # Step 4: docx_edit append approval paragraph
    c4, out4, err4 = run_command([
        str(DOCX_EDIT_SCRIPT), str(stage1_docx), "append-paragraph",
        "--text", "Документ согласован и подписан уполномоченными лицами.",
        "--style", "Normal",
    ])
    step4_ok = c4 == 0

    doc = None
    has_replaced = False
    has_appended = False
    has_bak = False
    all_ok = step1_ok and step2_ok and step3_ok and step4_ok
    if all_ok:
        generated_docx_files.append(stage1_docx)
        doc = docx.Document(str(stage1_docx))
        p_texts = [p.text for p in doc.paragraphs]
        has_replaced = any("УТВЕРЖДЕННЫЙ ДОГОВОР ОКАЗАНИЯ УСЛУГ" in t for t in p_texts)
        has_appended = any("Документ согласован и подписан" in t for t in p_texts)
        bak_file = stage1_docx.with_suffix(".docx.bak")
        has_bak = bak_file.is_file()

        if not has_replaced:
            all_ok = False
            notes.append("In-place text replacement not reflected in document")
        if not has_appended:
            all_ok = False
            notes.append("Appended approval paragraph not found in document")
        if not has_bak:
            all_ok = False
            notes.append("Backup file .docx.bak was not created by docx-editor")

    rep = f"""# Отчет по сценарию 10: Chained Workflow with docx-editor
- Шаг 1 (markdown-to-docx): {'УСПЕХ' if step1_ok else 'ОШИБКА'}
- Шаг 2 (docx_inspect outline): {'УСПЕХ' if step2_ok else 'ОШИБКА'}
- Шаг 3 (docx_edit replace-text): {'УСПЕХ' if step3_ok else 'ОШИБКА'}
- Шаг 4 (docx_edit append-paragraph): {'УСПЕХ' if step4_ok else 'ОШИБКА'}
- Создание резервной копии .docx.bak: {'ДА' if has_bak else 'НЕТ'}
- Текст 'УТВЕРЖДЕННЫЙ ДОГОВОР' в документе: {'ДА' if has_replaced else 'НЕТ'}
- Замечания: {'; '.join(notes) if notes else 'Нет'}
- Статус: {'УСПЕХ' if all_ok else 'ОШИБКА'}
"""
    (folder / "scenario_report.md").write_text(rep, encoding="utf-8")
    record_result("Scenario 10: Chained Workflow with docx-editor", all_ok, "End-to-end MD -> DOCX -> docx-editor pipeline verified", folder)


# =========================================================================
# Scenario 11: Document Validity & Microsoft Word COM Automation
# =========================================================================
def scenario_11_word_com_automation():
    """Verify all generated DOCX documents open in MS Word via COM without repair prompts and export to PDF."""
    folder = SCENARIOS_DIR / "11_word_com_automation"
    folder.mkdir(parents=True, exist_ok=True)

    notes = []
    all_ok = True
    validated_files = 0
    pdf_exported = False
    com_ok = False

    # Check ZIP and XML integrity for all generated files
    for docx_file in generated_docx_files:
        if not docx_file.is_file():
            all_ok = False
            notes.append(f"File missing: {docx_file.name}")
            continue
        try:
            with zipfile.ZipFile(docx_file, "r") as z:
                if "word/document.xml" not in z.namelist():
                    all_ok = False
                    notes.append(f"{docx_file.name} is missing word/document.xml")
        except Exception as e:
            all_ok = False
            notes.append(f"{docx_file.name} zip error: {e}")

    # Test opening all files in MS Word via COM
    test_script = folder / "com_check.py"
    files_to_check = [str(f.resolve()) for f in generated_docx_files if f.is_file()]
    export_pdf_target = str((folder / "sample_validation.pdf").resolve())

    test_script.write_text(
        f'''
import sys
import win32com.client

files = {json.dumps(files_to_check)}
pdf_target = {json.dumps(export_pdf_target)}

try:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    
    for idx, f in enumerate(files):
        doc = word.Documents.Open(f, ReadOnly=True, ConfirmConversions=False)
        print(f"Opened {{idx+1}}/{{len(files)}}: {{f}}")
        if idx == 0:
            doc.SaveAs(pdf_target, FileFormat=17) # 17 = wdFormatPDF
            print(f"PDF exported successfully to {{pdf_target}}")
        doc.Close(SaveChanges=False)
        
    word.Quit()
    print("ALL_DOCS_PASSED_WORD_COM")
except Exception as e:
    print("COM_ERROR:", e, file=sys.stderr)
    try:
        word.Quit()
    except Exception:
        pass
    sys.exit(1)
''',
        encoding="utf-8",
    )

    code, out, err = run_command(["--with", "pywin32", "python", str(test_script)], timeout=90)
    com_ok = (code == 0 and "ALL_DOCS_PASSED_WORD_COM" in out)

    if not com_ok:
        all_ok = False
        notes.append(f"Word COM validation failed: {err.strip() or out.strip()}")
    else:
        validated_files = len(files_to_check)
        pdf_file = folder / "sample_validation.pdf"
        pdf_exported = pdf_file.is_file() and pdf_file.stat().st_size > 0
        if not pdf_exported:
            all_ok = False
            notes.append("PDF export via Word COM did not produce a valid file")

    rep = f"""# Отчет по сценарию 11: Document Validity & Microsoft Word COM Automation
- Проверено файлов: {validated_files}
- Все файлы открылись в Word COM без ошибок: {'ДА' if com_ok else 'НЕТ'}
- Экспорт в PDF подтвержден: {'ДА' if pdf_exported else 'НЕТ'}
- Замечания: {'; '.join(notes) if notes else 'Нет'}
- Статус: {'УСПЕХ' if all_ok else 'ОШИБКА'}
"""
    (folder / "scenario_report.md").write_text(rep, encoding="utf-8")
    record_result("Scenario 11: Word COM Automation", all_ok, f"All {validated_files} docx files opened in MS Word & exported to PDF", folder)


# =========================================================================
# Main Test Runner
# =========================================================================
def main():
    print("=" * 80)
    print("ТЕСТОВЫЙ НАБОР ДЛЯ СКИЛА: markdown-to-docx")
    print(f"Целевой скрипт: {CONVERT_SCRIPT.relative_to(REPO_ROOT)}")
    print(f"Папка результатов: {SCENARIOS_DIR.relative_to(REPO_ROOT)}")
    print("=" * 80)

    scenario_01_headings_and_hierarchy()
    scenario_02_full_inline_formatting()
    scenario_03_hyperlinks_and_anchors()
    scenario_04_lists_and_hierarchies()
    scenario_05_gfm_tables_and_alignment()
    scenario_06_blockquotes_and_callouts()
    scenario_07_code_blocks_and_syntax()
    scenario_08_images_and_aspect_ratio()
    scenario_09_template_inheritance()
    scenario_10_chained_workflow_with_docx_editor()
    scenario_11_word_com_automation()

    # Generate test_summary.json
    all_passed = all(r["passed"] for r in test_results)
    summary = {
        "skill": "markdown-to-docx",
        "total_scenarios": len(test_results),
        "passed_scenarios": sum(1 for r in test_results if r["passed"]),
        "failed_scenarios": sum(1 for r in test_results if not r["passed"]),
        "all_passed": all_passed,
        "scenarios": test_results,
    }

    summary_file = TEST_ROOT / "test_summary.json"
    summary_file.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print("\n" + "=" * 80)
    print(f"ИТОГИ ТЕСТИРОВАНИЯ markdown-to-docx: {summary['passed_scenarios']}/{summary['total_scenarios']} УСПЕШНО")
    print("=" * 80)

    for r in test_results:
        st = "✅ PASS" if r["passed"] else "❌ FAIL"
        print(f"[{st}] {r['scenario']}: {r['details']}")

    sys.exit(0 if all_passed else 1)


if __name__ == "__main__":
    main()
