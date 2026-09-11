# Тестовая инфраструктура плагинов Antigravity

В этой директории собраны автоматизированные наборы тестов для проверки стабильности и качества работы скилов и плагинов Antigravity.

---

## Архитектура тестов

Тесты изолированы от продуктового кода плагинов (`.agents/plugins/`), чтобы плагины оставались портативными и легковесными при установке.

```text
tests/
├── run_all.py                  # Главный супер-раннер тестов
├── docx-tools/                 # Тесты для плагина docx-tools
│   ├── docx-editor/            # 16 сценариев комплексного тестирования docx-editor
│   │   ├── run_tests.py        # Автономный раннер docx-editor
│   │   ├── README_MANUAL_VERIFICATION.md
│   │   ├── test_summary.json
│   │   └── scenarios/
│   ├── docx-to-markdown/       # Тесты конвертации в Markdown
│   ├── docx-to-pdf/            # Тесты нативного экспорта в PDF
│   └── markdown-to-docx/       # Тесты компиляции DOCX из Markdown
└── excel-tools/                # Тесты будущего плагина Excel
```

---

## Запуск тестов

Для запуска требуется установленный пакетный менеджер [`uv`](https://docs.astral.sh/uv/):

### 1. Запуск всех тестов проекта
```bash
uv run tests/run_all.py
```

### 2. Запуск тестов конкретного плагина
```bash
uv run tests/run_all.py --plugin docx-tools
```

### 3. Автономный запуск тестов отдельного скила (например, docx-editor)
```bash
uv run tests/docx-tools/docx-editor/run_tests.py
```

Все зависимости (`python-docx`, `pywin32`, `pymupdf`) устанавливаются и изолируются автоматически виртуальным окружением `uv`.
