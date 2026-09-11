# Тестовый набор для скила `docx-to-markdown`

Комплексный набор автоматизированных тестов для скила конвертации документов Microsoft Word (`.docx`) в структурированный Markdown (`.md`) с выгрузкой изображений, форматированием таблиц и извлечением комментариев OpenXML.

---

## Назначение

Набор тестов обеспечивает верификацию скрипта [`docx_to_md.py`](file:///c:/Users/kekss/Project/Antigravity_IDE_configs/.agents/plugins/docx-tools/skills/docx-to-markdown/scripts/docx_to_md.py), предотвращает регрессии и гарантирует строгое соблюдение всех требований к конвертации:
1. **Экспорт изображений:** извлечение всех встроенных картинок в указанную папку (`--media-dir`) и замена ссылок в Markdown на относительные пути (без использования inline `data:image/...;base64`).
2. **Таблицы:** корректный рендеринг таблиц Word в стандартные таблицы GitHub Flavored Markdown (GFM) с разделителями `| --- |`.
3. **Комментарии Word:** извлечение комментариев из OpenXML (`word/comments.xml`) с автором, датой, цитатой целевого фрагмента документа и текстом комментария в блоке `## Комментарии / Comments`.
4. **Списки, стили и гиперссылки:** сохранение уровней заголовков H1..H6 (включая русские стили), списков, жирного/курсивного начертания и внешних ссылок.
5. **Колонтитулы и движки:** поддержка флагов `--include-headers`, `--no-comments`, выбор движков `markitdown`, `mammoth` и `auto`.

---

## Структура каталога

```
tests/docx-tools/docx-to-markdown/
├── README.md               # Документация тестового набора
├── run_tests.py            # Главный раннер тестов (PEP 723)
├── test_summary.json       # Сводка результатов последнего прогона
└── scenarios/              # Артефакты и отчёты по каждому сценарию
    ├── 01_headings_and_structure/
    ├── 02_tables_formatting/
    ├── 03_image_extraction_and_paths/
    ├── 04_comments_extraction/
    ├── 05_lists_styles_links/
    ├── 06_headers_footers/
    ├── 07_conversion_engines/
    ├── 08_real_world_samples/
    └── 09_cli_and_edge_cases/
```

---

## Описание тестовых сценариев

| Сценарий | Название | Что проверяется |
| :--- | :--- | :--- |
| **01** | `Headings & Structure` | Заголовки H1..H3, абзацы, русские стили («Заголовок 1..2»), сохранение иерархии. |
| **02** | `Tables to GFM` | Преобразование сложных и простых таблиц Word в GFM-таблицы с разделителями `\| --- \|` и сохранением данных ячеек. |
| **03** | `Images & Relative Paths` | Экспорт изображений на диск в `--media-dir`, проверка относительных ссылок `![alt](media/...)` и отсутствия inline base64. |
| **04** | `Comments Extraction` | Парсинг `word/comments.xml`, извлечение автора, даты, фрагмента текста и комментария; проверка флага `--no-comments`. |
| **05** | `Lists, Styles & Links` | Маркированные списки (`*`), жирный (`**bold**`), курсив (`*italic*`), гиперссылки `[text](url)`. |
| **06** | `Headers & Footers` | Работа флага `--include-headers` с генерацией блока цитаты `> **Колонтитул:**`. |
| **07** | `Conversion Engines` | Независимый запуск движков `--engine markitdown`, `--engine mammoth` и режима `auto`. |
| **08** | `Real Samples Suite` | Сквозной прогон на реальных файлах `file-sample_1.docx`, `file-sample_2.docx`, `file-sample_3.docx`. |
| **09** | `CLI & Edge Cases` | Обработка несуществующего файла (код 1), автосоздание вложенных путей, обработка минимального документа. |

---

## Запуск тестов

### 1. Автономный запуск с `uv` (рекомендуется):
```bash
uv run tests/docx-tools/docx-to-markdown/run_tests.py
```

### 2. Запуск через общий раннер репозитория:
```bash
uv run tests/run_all.py --plugin docx-tools
```
