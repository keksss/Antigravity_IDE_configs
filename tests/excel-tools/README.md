# Тестовый набор для плагина `excel-tools`

## Назначение
Тестирование скилов и скриптов плагина [`excel-tools`](../../.agents/plugins/excel-tools/) (работа с книгами Excel, формулами, стилями, универсальным пересчетом и инспекцией данных).

---

## Структура тестов

```text
tests/excel-tools/
├── README.md                         # Данное описание
└── xlsx/                             # Тестовый набор скилла xlsx
    ├── README_SCENARIOS.md           # Подробная матрица 16 тестовых сценариев
    ├── run_tests.py                  # Автоматизированный раннер всех сценариев
    ├── test_summary.json             # Автоматический отчет выполнения
    └── scenarios/                    # Каталоги и артефакты каждого сценария
```

---

## Запуск тестов

### 1. Запуск через общий супер-раннер Antigravity:
```bash
uv run tests/run_all.py --plugin excel-tools
```

### 2. Запуск напрямую раннера скилла `xlsx`:
```bash
uv run tests/excel-tools/xlsx/run_tests.py
```

### 3. Запуск конкретного сценария по фильтру:
```bash
uv run tests/excel-tools/xlsx/run_tests.py --filter 05
uv run tests/excel-tools/xlsx/run_tests.py --filter recalc
```
