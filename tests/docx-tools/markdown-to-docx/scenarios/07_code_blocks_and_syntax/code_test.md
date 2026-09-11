# Code Syntax Presentation

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
