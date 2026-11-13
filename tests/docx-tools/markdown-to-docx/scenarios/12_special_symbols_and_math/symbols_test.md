# Спецсимволы, математика и инлайн-HTML

Стрелки LaTeX: \rightarrow и $\rightarrow$, а также $\Rightarrow$, $\Leftrightarrow$, $\mapsto$, $\leftarrow$.
Формулы: $x_1 + y_2 \le 10 \pm 0.5$, $\alpha, \beta, \Delta, \Omega$, $\frac{1}{2}$, $\sqrt{x}$, $x \in \mathbb{R}$, $\forall \varepsilon > 0, \exists \delta \neq 0, \infty$.
Сложный индекс: $x_{total} + y^{max}$.

ASCII стрелки и типографика: A -> B, C --> D, E <-> F, G ==> H, I <=> J, точность +-5%, проверка a != b, ожидание..., права (c) (r) (tm).
Бюджет проекта составляет от $100 до $200 без наценок.

HTML сущности и теги: &rightarrow; и &rarr;, вода H<sub>2</sub>O, энергия E = mc<sup>2</sup>, <u>подчеркнутый</u>, <del>удаленный</del>, <mark>выделенный</mark>, клавиша <kbd>Ctrl+C</kbd>.
Эмодзи-шорткоды: :arrow_right: Готово :white_check_mark: Предупреждение :warning:

| Конструкция | Результат и перенос |
| :--- | :--- |
| Многострочная ячейка | Первая строка<br/>Вторая строка $\rightarrow$ итог |

Защита кода: `a -> b != c \rightarrow` не должна искажаться.

```python
def transform(x: int) -> bool:
    # \rightarrow and != must remain literal in code blocks
    return x != 0
```
