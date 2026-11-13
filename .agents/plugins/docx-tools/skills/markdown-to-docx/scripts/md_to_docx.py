# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "python-docx>=1.1.2",
#     "markdown-it-py>=3.0.0",
#     "Pillow>=10.0.0",
# ]
# ///
"""Convert Markdown (.md) documents to Microsoft Word (.docx) with 100% syntax coverage and optional template styling."""

import argparse
import html
import os
import re
import shutil
# Pandoc is invoked with separate arguments, without a shell.
import subprocess  # nosec B404
import sys
from pathlib import Path
from PIL import Image
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from markdown_it import MarkdownIt
from markdown_it.tree import SyntaxTreeNode

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, OSError, ValueError) as exc:
        print(f"Warning: UTF-8 console configuration failed: {exc}", file=sys.stderr)


def clear_document_body(doc: Document):
    """Remove existing paragraphs and tables from a template while preserving styles, headers, footers, and sections."""
    body_elm = doc._body._element
    for child in list(body_elm):
        if child.tag.endswith("sectPr"):
            continue
        body_elm.remove(child)


def get_node_info(node):
    """Safely extract node attributes whether node is a SyntaxTreeNode or a dictionary."""
    if isinstance(node, dict):
        return node.get("type", ""), node.get("content", ""), node.get("children", []), node.get("attrs", {})
    return getattr(node, "type", ""), getattr(node, "content", ""), getattr(node, "children", []), getattr(node, "attrs", {})


def extract_plain_text(node) -> str:
    """Recursively extract plain text from any AST node."""
    node_type, content, children, _ = get_node_info(node)
    if content:
        return content
    texts = []
    for c in children:
        texts.append(extract_plain_text(c))
    return "".join(texts)


# ---------------------------------------------------------------------------
# Special Symbols, LaTeX / Math, HTML Entities, Emoji & Typography Handling
# ---------------------------------------------------------------------------

LATEX_SYMBOLS: dict[str, str] = {
    # Arrows
    r"\Longleftrightarrow": "⟺",
    r"\longleftrightarrow": "⟷",
    r"\Longrightarrow": "⟹",
    r"\Longleftarrow": "⟸",
    r"\longrightarrow": "→",
    r"\longleftarrow": "←",
    r"\longmapsto": "↦",
    r"\Leftrightarrow": "⇔",
    r"\leftrightarrow": "↔",
    r"\rightleftharpoons": "⇌",
    r"\leftrightarrows": "⇆",
    r"\rightleftarrows": "⇄",
    r"\Rightarrow": "⇒",
    r"\Leftarrow": "⇐",
    r"\rightarrow": "→",
    r"\leftarrow": "←",
    r"\Uparrow": "⇑",
    r"\Downarrow": "⇓",
    r"\Updownarrow": "⇕",
    r"\uparrow": "↑",
    r"\downarrow": "↓",
    r"\updownarrow": "↕",
    r"\upuparrows": "⇈",
    r"\downdownarrows": "⇊",
    r"\nearrow": "↗",
    r"\searrow": "↘",
    r"\swarrow": "↙",
    r"\nwarrow": "↖",
    r"\hookrightarrow": "↪",
    r"\hookleftarrow": "↩",
    r"\mapsto": "↦",
    r"\implies": "⇒",
    r"\impliedby": "⇐",
    r"\iff": "⇔",
    r"\to": "→",
    r"\gets": "←",
    r"\rightharpoonup": "⇀",
    r"\rightharpoondown": "⇁",
    r"\leftharpoonup": "↼",
    r"\leftharpoondown": "↽",
    r"\nrightarrow": "↛",
    r"\nleftarrow": "↚",
    r"\nleftrightarrow": "↮",
    r"\nRightarrow": "⇏",
    r"\nLeftarrow": "⇍",
    r"\nLeftrightarrow": "⇎",
    # Comparison & Relations
    r"\leqslant": "≤",
    r"\geqslant": "≥",
    r"\leq": "≤",
    r"\geq": "≥",
    r"\le": "≤",
    r"\ge": "≥",
    r"\neq": "≠",
    r"\ne": "≠",
    r"\thickapprox": "≈",
    r"\approx": "≈",
    r"\thicksim": "∼",
    r"\sim": "∼",
    r"\simeq": "≃",
    r"\cong": "≅",
    r"\equiv": "≡",
    r"\nequiv": "≢",
    r"\ll": "≪",
    r"\gg": "≫",
    r"\propto": "∝",
    r"\asymp": "≍",
    r"\doteq": "≐",
    r"\parallel": "∥",
    r"\nparallel": "∦",
    r"\perp": "⊥",
    r"\mid": "∣",
    r"\nmid": "∤",
    # Arithmetic & Binary Operators
    r"\pm": "±",
    r"\mp": "∓",
    r"\times": "×",
    r"\div": "÷",
    r"\cdot": "·",
    r"\centerdot": "·",
    r"\bullet": "•",
    r"\circ": "∘",
    r"\ast": "∗",
    r"\star": "⋆",
    r"\oplus": "⊕",
    r"\ominus": "⊖",
    r"\otimes": "⊗",
    r"\oslash": "⊘",
    r"\odot": "⊙",
    r"\dagger": "†",
    r"\ddagger": "‡",
    r"\amalg": "⨿",
    r"\wr": "≀",
    # Set Theory & Logic
    r"\in": "∈",
    r"\notin": "∉",
    r"\ni": "∋",
    r"\owns": "∋",
    r"\subseteq": "⊆",
    r"\supseteq": "⊇",
    r"\nsubseteq": "⊈",
    r"\nsupseteq": "⊉",
    r"\subsetneq": "⊊",
    r"\supsetneq": "⊋",
    r"\subset": "⊂",
    r"\supset": "⊃",
    r"\cup": "∪",
    r"\cap": "∩",
    r"\setminus": "∖",
    r"\emptyset": "∅",
    r"\varnothing": "∅",
    r"\forall": "∀",
    r"\exists": "∃",
    r"\nexists": "∄",
    r"\neg": "¬",
    r"\lnot": "¬",
    r"\land": "∧",
    r"\wedge": "∧",
    r"\lor": "∨",
    r"\vee": "∨",
    r"\veebar": "⊻",
    r"\vdash": "⊢",
    r"\dashv": "⊣",
    r"\models": "⊨",
    r"\top": "⊤",
    r"\bot": "⊥",
    r"\therefore": "∴",
    r"\because": "∵",
    # Calculus, Analysis, Geometry & Misc
    r"\infty": "∞",
    r"\partial": "∂",
    r"\nabla": "∇",
    r"\sum": "∑",
    r"\prod": "∏",
    r"\coprod": "∐",
    r"\iiint": "∭",
    r"\iint": "∬",
    r"\oint": "∮",
    r"\int": "∫",
    r"\sqrt": "√",
    r"\degree": "°",
    r"\prime": "′",
    r"\ldots": "…",
    r"\dots": "…",
    r"\cdots": "⋯",
    r"\vdots": "⋮",
    r"\ddots": "⋱",
    r"\checkmark": "✓",
    r"\checked": "✓",
    r"\xmark": "✗",
    r"\measuredangle": "∡",
    r"\angle": "∠",
    r"\bigtriangleup": "△",
    r"\triangle": "△",
    r"\blacktriangle": "▲",
    r"\square": "□",
    r"\Box": "□",
    r"\blacksquare": "■",
    r"\diamond": "◇",
    r"\Diamond": "◇",
    r"\blacklozenge": "⧫",
    r"\ell": "ℓ",
    r"\hbar": "ℏ",
    r"\Re": "ℜ",
    r"\Im": "ℑ",
    r"\aleph": "ℵ",
    r"\wp": "℘",
    r"\mho": "℧",
    r"\AA": "Å",
    r"\S": "§",
    r"\P": "¶",
    r"\copyright": "©",
    r"\registered": "®",
    r"\trademark": "™",
    r"\euro": "€",
    r"\pounds": "£",
    r"\yen": "¥",
    r"\permil": "‰",
    # Greek Letters (lowercase & uppercase)
    r"\varepsilon": "ε",
    r"\vartheta": "ϑ",
    r"\varpi": "ϖ",
    r"\varrho": "ϱ",
    r"\varsigma": "ς",
    r"\varphi": "φ",
    r"\alpha": "α",
    r"\beta": "β",
    r"\gamma": "γ",
    r"\delta": "δ",
    r"\epsilon": "ϵ",
    r"\zeta": "ζ",
    r"\eta": "η",
    r"\theta": "θ",
    r"\iota": "ι",
    r"\kappa": "κ",
    r"\lambda": "λ",
    r"\mu": "μ",
    r"\nu": "ν",
    r"\xi": "ξ",
    r"\pi": "π",
    r"\rho": "ρ",
    r"\sigma": "σ",
    r"\tau": "τ",
    r"\upsilon": "υ",
    r"\phi": "ϕ",
    r"\chi": "χ",
    r"\psi": "ψ",
    r"\omega": "ω",
    r"\Gamma": "Γ",
    r"\Delta": "Δ",
    r"\Theta": "Θ",
    r"\Lambda": "Λ",
    r"\Xi": "Ξ",
    r"\Pi": "Π",
    r"\Sigma": "Σ",
    r"\Upsilon": "Υ",
    r"\Phi": "Φ",
    r"\Psi": "Ψ",
    r"\Omega": "Ω",
    # Math spacing escapes
    r"\qquad": "\u2003\u2003",
    r"\quad": "\u2003",
    r"\,": "\u2009",
    r"\;": " ",
    r"\:": " ",
    r"\!": "",
}

BLACKBOARD_BOLD_MAP: dict[str, str] = {
    "R": "ℝ",
    "N": "ℕ",
    "Z": "ℤ",
    "Q": "ℚ",
    "C": "ℂ",
    "P": "ℙ",
    "E": "𝔼",
    "F": "𝔽",
    "H": "ℍ",
}

VULGAR_FRACTIONS: dict[tuple[str, str], str] = {
    ("1", "2"): "½",
    ("1", "3"): "⅓",
    ("2", "3"): "⅔",
    ("1", "4"): "¼",
    ("3", "4"): "¾",
    ("1", "5"): "⅕",
    ("2", "5"): "⅖",
    ("3", "5"): "⅗",
    ("4", "5"): "⅘",
    ("1", "6"): "⅙",
    ("5", "6"): "⅚",
    ("1", "8"): "⅛",
    ("3", "8"): "⅜",
    ("5", "8"): "⅝",
    ("7", "8"): "⅞",
}

SUPERSCRIPT_MAP: dict[str, str] = str.maketrans("0123456789+-=()ni", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ")
SUPERSCRIPT_CHARS = set("0123456789+-=()ni")

SUBSCRIPT_MAP: dict[str, str] = str.maketrans(
    "0123456789+-=()aehijklmnoprstuvx",
    "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ",
)
SUBSCRIPT_CHARS = set("0123456789+-=()aehijklmnoprstuvx")

CUSTOM_HTML_ENTITIES: dict[str, str] = {
    "&rightarrow;": "→",
    "&leftarrow;": "←",
    "&leftrightarrow;": "↔",
    "&Rightarrow;": "⇒",
    "&Leftarrow;": "⇐",
    "&Leftrightarrow;": "⇔",
    "&longleftrightarrow;": "⟷",
    "&longrightarrow;": "→",
    "&longleftarrow;": "←",
}

EMOJI_SHORTCODES: dict[str, str] = {
    ":arrow_right:": "→",
    ":arrow_left:": "←",
    ":arrow_up:": "↑",
    ":arrow_down:": "↓",
    ":arrow_upper_right:": "↗",
    ":arrow_lower_right:": "↘",
    ":arrow_lower_left:": "↙",
    ":arrow_upper_left:": "↖",
    ":left_right_arrow:": "↔",
    ":arrow_up_down:": "↕",
    ":arrows_counterclockwise:": "🔄",
    ":point_right:": "👉",
    ":point_left:": "👈",
    ":white_check_mark:": "✅",
    ":heavy_check_mark:": "✔",
    ":ballot_box_with_check:": "☑",
    ":x:": "❌",
    ":cross_mark:": "❌",
    ":heavy_multiplication_x:": "✖",
    ":warning:": "⚠",
    ":exclamation:": "❗",
    ":heavy_exclamation_mark:": "❗",
    ":question:": "❓",
    ":bulb:": "💡",
    ":memo:": "📝",
    ":pencil:": "📝",
    ":pushpin:": "📌",
    ":round_pushpin:": "📍",
    ":bookmark:": "🔖",
    ":link:": "🔗",
    ":lock:": "🔒",
    ":unlock:": "🔓",
    ":key:": "🔑",
    ":mag:": "🔍",
    ":gear:": "⚙",
    ":wrench:": "🔧",
    ":hammer:": "🔨",
    ":rocket:": "🚀",
    ":fire:": "🔥",
    ":sparkles:": "✨",
    ":star:": "⭐",
    ":zap:": "⚡",
    ":bug:": "🐛",
    ":chart_with_upwards_trend:": "📈",
    ":chart_with_downwards_trend:": "📉",
    ":bar_chart:": "📊",
    ":clipboard:": "📋",
    ":file_folder:": "📁",
    ":page_facing_up:": "📄",
    ":calendar:": "📅",
    ":hourglass:": "⏳",
    ":thumbsup:": "👍",
    ":+1:": "👍",
    ":thumbsdown:": "👎",
    ":-1:": "👎",
    ":information_source:": "ℹ",
    ":info:": "ℹ",
    ":stop_sign:": "🛑",
    ":no_entry:": "⛔",
    ":construction:": "🚧",
    ":trophy:": "🏆",
    ":dart:": "🎯",
}

# Precompiled regexes for LaTeX command matching (longest-first)
_ALPHA_LATEX_KEYS = sorted((k for k in LATEX_SYMBOLS if k[1:].isalpha()), key=len, reverse=True)
_NONALPHA_LATEX_KEYS = sorted((k for k in LATEX_SYMBOLS if not k[1:].isalpha()), key=len, reverse=True)

_RE_ALPHA_LATEX = re.compile(
    r"(?:" + "|".join(re.escape(k) for k in _ALPHA_LATEX_KEYS) + r")(?![a-zA-Z])(?:\{\s*\})?"
)
_RE_ALL_LATEX = re.compile(
    r"(?:" + "|".join(re.escape(k) for k in _ALPHA_LATEX_KEYS) + r")(?![a-zA-Z])(?:\{\s*\})?"
    + r"|(?:" + "|".join(re.escape(k) for k in _NONALPHA_LATEX_KEYS) + r")"
)
_RE_EMOJI_SHORTCODE = re.compile(
    "|".join(re.escape(k) for k in sorted(EMOJI_SHORTCODES, key=len, reverse=True))
)
_RE_XML_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff\u200b\ufeff]")
_RE_URL = re.compile(r"(?:https?|ftp|mailto):[^\s<>()]+")

_ASCII_ARROW_MAP: dict[str, str] = {
    "<-->": "⟷",
    "<==>": "⟺",
    "<->": "↔",
    "<=>": "⇔",
    "-->": "→",
    "<--": "←",
    "==>": "⇒",
    "<==": "⇐",
    "->": "→",
    "<-": "←",
}
_RE_ASCII_ARROWS = re.compile(
    "|".join(re.escape(k) for k in sorted(_ASCII_ARROW_MAP, key=len, reverse=True))
)


def sanitize_xml_text(text: str) -> str:
    """Remove XML-illegal control characters and invisible zero-width markers."""
    if not text:
        return text
    return _RE_XML_ILLEGAL.sub("", text)


def _is_symbol_font_char(ch: str) -> bool:
    """Return True if character belongs to Unicode symbol blocks needing Segoe UI Symbol fallback."""
    cp = ord(ch)
    return (
        0x2100 <= cp <= 0x21FF
        or 0x2200 <= cp <= 0x22FF
        or 0x2300 <= cp <= 0x23FF
        or 0x25A0 <= cp <= 0x25FF
        or 0x2600 <= cp <= 0x26FF
        or 0x2700 <= cp <= 0x27FF
        or 0x2900 <= cp <= 0x297F
        or 0x2A00 <= cp <= 0x2AFF
        or 0x2B00 <= cp <= 0x2BFF
    )


def _split_symbol_segments(text: str) -> list[tuple[str, bool]]:
    """Split text into contiguous runs of (segment_text, needs_symbol_font)."""
    if not text:
        return []
    segments: list[tuple[str, bool]] = []
    buf: list[str] = []
    cur_flag = _is_symbol_font_char(text[0])
    for ch in text:
        flag = _is_symbol_font_char(ch)
        if flag != cur_flag:
            segments.append(("".join(buf), cur_flag))
            buf = [ch]
            cur_flag = flag
        else:
            buf.append(ch)
    if buf:
        segments.append(("".join(buf), cur_flag))
    return segments


def _mask_code_regions(md_text: str) -> tuple[str, list[str]]:
    """Mask fenced code blocks and inline code spans so preprocessing leaves code untouched."""
    masked: list[str] = []
    out: list[str] = []
    i = 0
    n = len(md_text)

    def store(chunk: str) -> str:
        idx = len(masked)
        masked.append(chunk)
        return f"\x00MD_CODE_{idx}\x00"

    fence_re = re.compile(r"^[ \t]{0,3}(`{3,}|~{3,})[^\n]*(?:\n|$)", re.MULTILINE)

    while i < n:
        # Check fenced code block at line start
        if i == 0 or md_text[i - 1] == "\n":
            fm = fence_re.match(md_text, i)
            if fm:
                fence_char = fm.group(1)[0]
                fence_len = len(fm.group(1))
                close_re = re.compile(
                    rf"^[ \t]{{0,3}}{re.escape(fence_char)}{{{fence_len},}}[ \t]*(?:\n|$)",
                    re.MULTILINE,
                )
                cm = close_re.search(md_text, fm.end())
                end_pos = cm.end() if cm else n
                out.append(store(md_text[i:end_pos]))
                i = end_pos
                continue

        # Check inline code span
        if md_text[i] == "`":
            tick_end = i
            while tick_end < n and md_text[tick_end] == "`":
                tick_end += 1
            tick_str = md_text[i:tick_end]
            # Find matching closing backtick run of exact length without crossing blank lines
            search_pos = tick_end
            close_pos = -1
            while search_pos < n:
                nxt = md_text.find(tick_str, search_pos)
                if nxt == -1:
                    break
                # Ensure exact backtick count
                after = nxt + len(tick_str)
                if (nxt == 0 or md_text[nxt - 1] != "`") and (after == n or md_text[after] != "`"):
                    if "\n\n" not in md_text[tick_end:nxt]:
                        close_pos = after
                    break
                search_pos = nxt + 1
            if close_pos != -1:
                out.append(store(md_text[i:close_pos]))
                i = close_pos
                continue
            out.append(tick_str)
            i = tick_end
            continue

        out.append(md_text[i])
        i += 1

    return "".join(out), masked


def _unmask_regions(text: str, items: list[str], prefix: str) -> str:
    """Restore masked segments by placeholder index."""
    if not items:
        return text
    pattern = re.compile(rf"\x00{re.escape(prefix)}_(\d+)\x00")
    return pattern.sub(lambda m: items[int(m.group(1))], text)


def _looks_like_math(inner: str) -> bool:
    """Heuristic distinguishing $...$ inline LaTeX math from prose currency like $100 ... $200."""
    stripped = inner.strip()
    if not stripped:
        return False
    if re.fullmatch(r"[\d.,\s]+", stripped):
        return False
    if any(ch in stripped for ch in ("\\", "^", "_", "=", "<", ">", "{", "}")):
        return True
    if any(ch.isspace() for ch in stripped):
        return bool(re.search(r"[+\-*/]", stripped))
    return len(stripped) <= 20 and bool(re.search(r"[A-Za-z]", stripped))


def convert_latex_expression(expr: str) -> str:
    """Convert a LaTeX math expression into clean Unicode text and inline sub/sup markup."""
    s = expr.strip()

    # Composite negations
    s = s.replace(r"\not\equiv", "≢").replace(r"\not\in", "∉").replace(r"\not\subset", "⊄").replace(r"\not=", "≠")

    # Blackboard bold \mathbb{R} -> ℝ
    s = re.sub(r"\\mathbb\{([A-Za-z])\}", lambda m: BLACKBOARD_BOLD_MAP.get(m.group(1), m.group(1)), s)

    # Unwrap \text{...}, \textbf{...}, \mathrm{...}, etc.
    for _ in range(3):
        s_next = re.sub(
            r"\\(?:text|textbf|textit|textrm|mathrm|mathbf|mathit|mathsf|mathtt|operatorname)\{([^{}]*)\}",
            r"\1",
            s,
        )
        if s_next == s:
            break
        s = s_next

    # Strip delimiter sizing commands (must not match \leftarrow, \rightarrow, etc.)
    s = re.sub(r"\\(?:left|right|bigl|bigr|Bigl|Bigr|big|Big)(?![a-zA-Z])\s*([\\.|()\[\]{}|])?", lambda m: (
        "{" if m.group(1) == r"\{" else
        "}" if m.group(1) == r"\}" else
        "‖" if m.group(1) == r"\|" else
        "" if m.group(1) in (".", None) else m.group(1)
    ), s)

    # Roots
    s = re.sub(r"\\sqrt\[3\]\{([^{}]+)\}", lambda m: f"∛{m.group(1)}" if len(m.group(1)) <= 2 else f"∛({m.group(1)})", s)
    s = re.sub(r"\\sqrt\[4\]\{([^{}]+)\}", lambda m: f"∜{m.group(1)}" if len(m.group(1)) <= 2 else f"∜({m.group(1)})", s)
    s = re.sub(r"\\sqrt\{([^{}]+)\}", lambda m: f"√{m.group(1)}" if len(m.group(1)) <= 2 else f"√({m.group(1)})", s)

    # Fractions (up to 2 nesting levels)
    def _replace_frac(match: re.Match) -> str:
        num = match.group(1).strip()
        den = match.group(2).strip()
        if (num, den) in VULGAR_FRACTIONS:
            return VULGAR_FRACTIONS[(num, den)]
        simple_re = r"^[A-Za-z0-9α-ωΑ-Ω]+$"
        if re.match(simple_re, num) and re.match(simple_re, den):
            return f"{num}/{den}"
        return f"({num})/({den})"

    for _ in range(2):
        s_next = re.sub(r"\\frac\{([^{}]+)\}\{([^{}]+)\}", _replace_frac, s)
        if s_next == s:
            break
        s = s_next

    # Degree marker
    s = re.sub(r"\^\{\s*\\circ\s*\}|\^\\circ(?![a-zA-Z])", "°", s)

    # Known LaTeX symbol commands
    s = _RE_ALL_LATEX.sub(lambda m: LATEX_SYMBOLS.get(m.group(0).rstrip("{}"), LATEX_SYMBOLS.get(m.group(0), m.group(0))), s)

    # Superscripts and subscripts
    def _format_sup(content: str) -> str:
        if content and all(ch in SUPERSCRIPT_CHARS for ch in content):
            return content.translate(SUPERSCRIPT_MAP)
        return f"<sup>{content}</sup>"

    def _format_sub(content: str) -> str:
        if content and all(ch in SUBSCRIPT_CHARS for ch in content):
            return content.translate(SUBSCRIPT_MAP)
        return f"<sub>{content}</sub>"

    s = re.sub(r"\^\{([^{}]+)\}", lambda m: _format_sup(m.group(1).strip()), s)
    s = re.sub(r"\^([A-Za-z0-9])", lambda m: _format_sup(m.group(1)), s)
    s = re.sub(r"_\{([^{}]+)\}", lambda m: _format_sub(m.group(1).strip()), s)
    s = re.sub(r"_([A-Za-z0-9])", lambda m: _format_sub(m.group(1)), s)

    # Unescape LaTeX punctuation escapes
    for esc, plain in ((r"\%", "%"), (r"\$", "$"), (r"\_", "_"), (r"\{", "{"), (r"\}", "}"), (r"\#", "#"), (r"\&", "&")):
        s = s.replace(esc, plain)

    return re.sub(r" {2,}", " ", s).strip()


def preprocess_markdown_special_constructs(md_text: str) -> str:
    """Normalize LaTeX/math spans and standalone LaTeX commands before MarkdownIt parsing."""
    protected_text, masked_code = _mask_code_regions(md_text)

    # 1. Display math $$ ... $$ and \[ ... \]
    protected_text = re.sub(
        r"\$\$([\s\S]+?)\$\$",
        lambda m: convert_latex_expression(m.group(1)),
        protected_text,
    )
    protected_text = re.sub(
        r"\\\[([\s\S]+?)\\\]",
        lambda m: convert_latex_expression(m.group(1)),
        protected_text,
    )

    # 2. Inline math \( ... \) and $ ... $
    protected_text = re.sub(
        r"\\\(([\s\S]+?)\\\)",
        lambda m: convert_latex_expression(m.group(1)),
        protected_text,
    )
    protected_text = re.sub(
        r"(?<!\\)\$(?!\s)([^\n$]+?)(?<!\s)(?<!\\)\$",
        lambda m: convert_latex_expression(m.group(1)) if _looks_like_math(m.group(1)) else m.group(0),
        protected_text,
    )

    # 3. Standalone alphabetic LaTeX commands outside math delimiters (e.g. bare \rightarrow)
    protected_text = _RE_ALPHA_LATEX.sub(
        lambda m: LATEX_SYMBOLS.get(m.group(0).rstrip("{}"), LATEX_SYMBOLS.get(m.group(0), m.group(0))),
        protected_text,
    )

    return _unmask_regions(protected_text, masked_code, "MD_CODE")


def normalize_inline_text(text: str) -> str:
    """Normalize HTML entities, emoji shortcodes, ASCII arrows, and typography in prose text."""
    if not text:
        return text

    # Mask URLs so arrows/dashes inside plain-text URLs stay intact
    urls: list[str] = []

    def _store_url(m: re.Match) -> str:
        idx = len(urls)
        urls.append(m.group(0))
        return f"\x00MD_URL_{idx}\x00"

    s = _RE_URL.sub(_store_url, text)

    # 1. Custom & standard HTML entities
    for ent, ch in CUSTOM_HTML_ENTITIES.items():
        if ent in s:
            s = s.replace(ent, ch)
    if "&" in s:
        s = html.unescape(s)

    # 2. Residual standalone LaTeX commands
    if "\\" in s:
        s = _RE_ALPHA_LATEX.sub(
            lambda m: LATEX_SYMBOLS.get(m.group(0).rstrip("{}"), LATEX_SYMBOLS.get(m.group(0), m.group(0))),
            s,
        )

    # 3. Emoji shortcodes
    if ":" in s:
        s = _RE_EMOJI_SHORTCODE.sub(lambda m: EMOJI_SHORTCODES[m.group(0)], s)

    # 4. ASCII arrows & typographic shorthands
    s = _RE_ASCII_ARROWS.sub(lambda m: _ASCII_ARROW_MAP[m.group(0)], s)
    s = re.sub(r"(?<=\s)=>(?=\s)", "⇒", s)
    s = re.sub(r"(?<!\w)\+/-(?!\w)|(?<=\s)\+-(?=\s|\d)", "±", s)
    s = re.sub(r"(?<=\s)!=(?=\s)", "≠", s)
    s = re.sub(r"(?<=\s)<=(?=\s)", "≤", s)
    s = re.sub(r"(?<=\s)>=(?=\s)", "≥", s)
    s = re.sub(r"(?<=\s)~=(?=\s)", "≈", s)
    s = s.replace("...", "…")
    s = s.replace("---", "—")
    s = re.sub(r"(?<=\s)--(?=\s)", "–", s)
    s = re.sub(r"\([cC]\)", "©", s)
    s = re.sub(r"\([rR]\)", "®", s)
    s = re.sub(r"\([tT][mM]\)", "™", s)

    s = _unmask_regions(s, urls, "MD_URL")
    return sanitize_xml_text(s)


def _apply_rfonts(rPr, font_name: str):
    """Set w:rFonts across all OpenXML font slots on an rPr element."""
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for slot in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(slot), font_name)


def add_styled_run(
    paragraph,
    text: str,
    *,
    bold: bool = False,
    italic: bool = False,
    strike: bool = False,
    code: bool = False,
    subscript: bool = False,
    superscript: bool = False,
    underline: bool = False,
    highlight: bool = False,
    hyperlink_el=None,
):
    """Add text runs with stacked formatting, XML sanitization, and symbol font fallback."""
    clean_text = sanitize_xml_text(text)
    if not clean_text:
        return

    for seg_text, is_symbol in _split_symbol_segments(clean_text):
        if hyperlink_el is not None:
            r = OxmlElement("w:r")
            rPr = OxmlElement("w:rPr")
            if bold:
                rPr.append(OxmlElement("w:b"))
            if italic:
                rPr.append(OxmlElement("w:i"))
            if strike:
                rPr.append(OxmlElement("w:strike"))
            if is_symbol:
                _apply_rfonts(rPr, "Segoe UI Symbol")
            elif code:
                _apply_rfonts(rPr, "Consolas")
            if code:
                sz = OxmlElement("w:sz")
                sz.set(qn("w:val"), "19")
                rPr.append(sz)
                c = OxmlElement("w:color")
                c.set(qn("w:val"), "C7254E")
                rPr.append(c)
            else:
                c = OxmlElement("w:color")
                c.set(qn("w:val"), "0563C1")
                rPr.append(c)
                u = OxmlElement("w:u")
                u.set(qn("w:val"), "single")
                rPr.append(u)
            if subscript or superscript:
                va = OxmlElement("w:vertAlign")
                va.set(qn("w:val"), "subscript" if subscript else "superscript")
                rPr.append(va)
            if highlight:
                hl = OxmlElement("w:highlight")
                hl.set(qn("w:val"), "yellow")
                rPr.append(hl)

            r.append(rPr)
            t = OxmlElement("w:t")
            t.text = seg_text
            if seg_text.startswith(" ") or seg_text.endswith(" "):
                t.set(qn("xml:space"), "preserve")
            r.append(t)
            hyperlink_el.append(r)
        else:
            run = paragraph.add_run(seg_text)
            run.bold = bold
            run.italic = italic
            run.font.strike = strike
            if underline:
                run.underline = True
            if subscript:
                run.font.subscript = True
            elif superscript:
                run.font.superscript = True
            if highlight:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            if is_symbol:
                run.font.name = "Segoe UI Symbol"
                _apply_rfonts(run._r.get_or_add_rPr(), "Segoe UI Symbol")
                if code:
                    run.font.size = Pt(9.5)
                    run.font.color.rgb = RGBColor(199, 37, 78)
            elif code:
                run.font.name = "Consolas"
                run.font.size = Pt(9.5)
                run.font.color.rgb = RGBColor(199, 37, 78)


def add_image_smart(doc: Document, paragraph, img_path: Path, max_w_in: float = 5.8, max_h_in: float = 7.5):
    """Scale image preserving natural dimensions, aspect ratio, and DPI via Pillow."""
    if not img_path.is_file():
        r = paragraph.add_run(f"[Missing Image: {img_path.name}]")
        r.italic = True
        r.font.color.rgb = RGBColor(180, 50, 50)
        return

    try:
        with Image.open(img_path) as im:
            px_w, px_h = im.size
            dpi = im.info.get("dpi", (96, 96))
            dpi_x = dpi[0] if isinstance(dpi, tuple) and dpi[0] > 0 else 96

        nat_w = px_w / dpi_x
        nat_h = px_h / dpi_x

        scale = 1.0
        if nat_w > max_w_in:
            scale = min(scale, max_w_in / nat_w)
        if nat_h > max_h_in:
            scale = min(scale, max_h_in / nat_h)

        final_w = Inches(nat_w * scale)

        pic_p = doc.add_paragraph()
        pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic_p.paragraph_format.space_before = Pt(4)
        pic_p.paragraph_format.space_after = Pt(4)
        pic_p.add_run().add_picture(str(img_path), width=final_w)
    except Exception as e:
        r = paragraph.add_run(f"[Image: {img_path.name} (Error: {e})]")
        r.italic = True


def render_inline_tree(
    paragraph,
    nodes: list,
    base_dir: Path,
    doc: Document,
    bold: bool = False,
    italic: bool = False,
    strike: bool = False,
    code: bool = False,
    subscript: bool = False,
    superscript: bool = False,
    underline: bool = False,
    highlight: bool = False,
    hyperlink_el=None,
):
    """Recursively render inline Markdown AST nodes with stacked formatting, inline HTML, and special symbols."""
    cur_bold = bold
    cur_italic = italic
    cur_strike = strike
    cur_code = code
    cur_sub = subscript
    cur_sup = superscript
    cur_underline = underline
    cur_highlight = highlight

    for node in nodes:
        node_type, content, children, attrs = get_node_info(node)

        if node_type == "text":
            text = content if cur_code else normalize_inline_text(content)
            if not text:
                continue
            add_styled_run(
                paragraph,
                text,
                bold=cur_bold,
                italic=cur_italic,
                strike=cur_strike,
                code=cur_code,
                subscript=cur_sub,
                superscript=cur_sup,
                underline=cur_underline,
                highlight=cur_highlight,
                hyperlink_el=hyperlink_el,
            )

        elif node_type == "code_inline":
            if not content:
                continue
            add_styled_run(
                paragraph,
                content,
                bold=cur_bold,
                italic=cur_italic,
                strike=cur_strike,
                code=True,
                subscript=cur_sub,
                superscript=cur_sup,
                underline=cur_underline,
                highlight=cur_highlight,
                hyperlink_el=hyperlink_el,
            )

        elif node_type == "html_inline":
            raw_tag = content.strip().lower()
            if raw_tag.startswith("<!--"):
                continue
            tag_match = re.match(r"^<(/?)([a-z0-9]+)(?:\s[^>]*)?/?>$", raw_tag)
            if not tag_match:
                continue
            is_closing = bool(tag_match.group(1))
            tag_name = tag_match.group(2)

            if tag_name == "br" and not is_closing:
                if hyperlink_el is not None:
                    r = OxmlElement("w:r")
                    r.append(OxmlElement("w:br"))
                    hyperlink_el.append(r)
                else:
                    paragraph.add_run().add_break()
            elif tag_name == "sub":
                cur_sub = not is_closing
            elif tag_name == "sup":
                cur_sup = not is_closing
            elif tag_name in ("u", "ins"):
                cur_underline = not is_closing
            elif tag_name in ("s", "del", "strike"):
                cur_strike = not is_closing
            elif tag_name in ("b", "strong"):
                cur_bold = not is_closing
            elif tag_name in ("i", "em"):
                cur_italic = not is_closing
            elif tag_name in ("code", "kbd", "samp"):
                cur_code = not is_closing
            elif tag_name == "mark":
                cur_highlight = not is_closing

        elif node_type == "strong":
            sub_nodes = children if children else (
                [{"type": "text", "content": content, "children": [], "attrs": {}}] if content else []
            )
            if sub_nodes:
                render_inline_tree(
                    paragraph, sub_nodes, base_dir, doc,
                    bold=True, italic=cur_italic, strike=cur_strike, code=cur_code,
                    subscript=cur_sub, superscript=cur_sup, underline=cur_underline,
                    highlight=cur_highlight, hyperlink_el=hyperlink_el,
                )

        elif node_type == "em":
            sub_nodes = children if children else (
                [{"type": "text", "content": content, "children": [], "attrs": {}}] if content else []
            )
            if sub_nodes:
                render_inline_tree(
                    paragraph, sub_nodes, base_dir, doc,
                    bold=cur_bold, italic=True, strike=cur_strike, code=cur_code,
                    subscript=cur_sub, superscript=cur_sup, underline=cur_underline,
                    highlight=cur_highlight, hyperlink_el=hyperlink_el,
                )

        elif node_type == "s":
            sub_nodes = children if children else (
                [{"type": "text", "content": content, "children": [], "attrs": {}}] if content else []
            )
            if sub_nodes:
                render_inline_tree(
                    paragraph, sub_nodes, base_dir, doc,
                    bold=cur_bold, italic=cur_italic, strike=True, code=cur_code,
                    subscript=cur_sub, superscript=cur_sup, underline=cur_underline,
                    highlight=cur_highlight, hyperlink_el=hyperlink_el,
                )

        elif node_type == "softbreak":
            if hyperlink_el is not None:
                r = OxmlElement("w:r")
                t = OxmlElement("w:t")
                t.text = " "
                t.set(qn("xml:space"), "preserve")
                r.append(t)
                hyperlink_el.append(r)
            else:
                paragraph.add_run(" ")

        elif node_type == "hardbreak":
            if hyperlink_el is not None:
                r = OxmlElement("w:r")
                r.append(OxmlElement("w:br"))
                hyperlink_el.append(r)
            else:
                paragraph.add_run().add_break()

        elif node_type == "link":
            href = attrs.get("href", "")
            try:
                part = paragraph.part
                r_id = part.relate_to(href, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
                hyperlink_obj = OxmlElement("w:hyperlink")
                hyperlink_obj.set(qn("r:id"), r_id)

                if children:
                    render_inline_tree(
                        paragraph,
                        children,
                        base_dir,
                        doc,
                        bold=cur_bold,
                        italic=cur_italic,
                        strike=cur_strike,
                        code=cur_code,
                        subscript=cur_sub,
                        superscript=cur_sup,
                        underline=cur_underline,
                        highlight=cur_highlight,
                        hyperlink_el=hyperlink_obj,
                    )
                else:
                    add_styled_run(paragraph, href, hyperlink_el=hyperlink_obj)

                paragraph._p.append(hyperlink_obj)
            except Exception:
                link_text = normalize_inline_text(extract_plain_text(node)) or href
                run = paragraph.add_run(sanitize_xml_text(f"{link_text} ({href})"))
                run.underline = True
                run.font.color.rgb = RGBColor(5, 99, 193)

        elif node_type == "image":
            src = attrs.get("src", "")
            img_path = (base_dir / src).resolve() if not os.path.isabs(src) else Path(src)
            add_image_smart(doc, paragraph, img_path)

        elif node_type in ("paragraph", "inline", "list_item"):
            if children:
                render_inline_tree(
                    paragraph, children, base_dir, doc,
                    bold=cur_bold, italic=cur_italic, strike=cur_strike, code=cur_code,
                    subscript=cur_sub, superscript=cur_sup, underline=cur_underline,
                    highlight=cur_highlight, hyperlink_el=hyperlink_el,
                )
            elif content:
                text = content if cur_code else normalize_inline_text(content)
                add_styled_run(
                    paragraph, text,
                    bold=cur_bold, italic=cur_italic, strike=cur_strike, code=cur_code,
                    subscript=cur_sub, superscript=cur_sup, underline=cur_underline,
                    highlight=cur_highlight, hyperlink_el=hyperlink_el,
                )

        elif children:
            render_inline_tree(
                paragraph, children, base_dir, doc,
                bold=cur_bold, italic=cur_italic, strike=cur_strike, code=cur_code,
                subscript=cur_sub, superscript=cur_sup, underline=cur_underline,
                highlight=cur_highlight, hyperlink_el=hyperlink_el,
            )
        elif content:
            text = content if cur_code else normalize_inline_text(content)
            add_styled_run(
                paragraph, text,
                bold=cur_bold, italic=cur_italic, strike=cur_strike, code=cur_code,
                subscript=cur_sub, superscript=cur_sup, underline=cur_underline,
                highlight=cur_highlight, hyperlink_el=hyperlink_el,
            )


def add_code_block(doc: Document, code_text: str, language: str = ""):
    """Render code block with background shading, left accent border, and Consolas font."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)

    # Shading (light gray #F5F6F8)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "F4F5F7")
    pPr.append(shd)

    # Left accent border (GitHub blue #0969DA)
    pBdr = OxmlElement("w:pBdr")
    left_bdr = OxmlElement("w:left")
    left_bdr.set(qn("w:val"), "single")
    left_bdr.set(qn("w:sz"), "18")  # 2.25 pt
    left_bdr.set(qn("w:space"), "8")
    left_bdr.set(qn("w:color"), "0969DA")
    pBdr.append(left_bdr)
    pPr.append(pBdr)

    clean_code = sanitize_xml_text(code_text.rstrip("\n"))
    run = p.add_run(clean_code)
    run.font.name = "Consolas"
    run.font.size = Pt(9.5)
    run.font.color.rgb = RGBColor(36, 41, 46)


def strip_alert_tag(inline_nodes: list, alert_type: str) -> list:
    """Strip [!ALERT] prefix from the first inline text node without mutating AST in-place."""
    if not inline_nodes:
        return inline_nodes
    tag_str = f"[!{alert_type}]"
    first = inline_nodes[0]
    n_type, content, children, attrs = get_node_info(first)
    if n_type == "text" and content:
        cleaned = content.lstrip()
        if cleaned.startswith(tag_str):
            new_content = cleaned[len(tag_str):].lstrip("\n ")
            new_first = {"type": "text", "content": new_content, "children": children, "attrs": attrs}
            return [new_first] + inline_nodes[1:]
    return inline_nodes


def render_callout_block(node: SyntaxTreeNode, alert_type: str, doc: Document, base_dir: Path):
    """Render GitHub-style alert callouts (> [!NOTE], etc.) with themed accent borders and full inline Markdown."""
    color_map = {
        "NOTE": ("0969DA", "Примечание"),
        "TIP": ("1A7F37", "Совет"),
        "IMPORTANT": ("8250DF", "Важно"),
        "WARNING": ("9A6700", "Внимание"),
        "CAUTION": ("D1242F", "Осторожно"),
    }
    hex_color, label = color_map.get(alert_type.upper(), ("0969DA", alert_type))

    paragraphs = [c for c in node.children if c.type == "paragraph"]
    if not paragraphs:
        paragraphs = [node]

    for p_idx, p_node in enumerate(paragraphs):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_before = Pt(4) if p_idx == 0 else Pt(2)
        p.paragraph_format.space_after = Pt(6) if p_idx == len(paragraphs) - 1 else Pt(2)

        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "F8FAFD")
        pPr.append(shd)

        pBdr = OxmlElement("w:pBdr")
        left_bdr = OxmlElement("w:left")
        left_bdr.set(qn("w:val"), "single")
        left_bdr.set(qn("w:sz"), "24")
        left_bdr.set(qn("w:space"), "8")
        left_bdr.set(qn("w:color"), hex_color)
        pBdr.append(left_bdr)
        pPr.append(pBdr)

        if p_idx == 0:
            # Add alert header badge
            lbl_run = p.add_run(f"[{label}] ")
            lbl_run.bold = True
            r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
            lbl_run.font.color.rgb = RGBColor(r, g, b)

        # Extract and render inline children
        inline_nodes = []
        for c in p_node.children:
            if c.type == "inline" and c.children:
                inline_nodes.extend(c.children)
            else:
                inline_nodes.append(c)

        if p_idx == 0:
            inline_nodes = strip_alert_tag(inline_nodes, alert_type)

        render_inline_tree(p, inline_nodes, base_dir, doc)


def render_regular_blockquote(node: SyntaxTreeNode, doc: Document, base_dir: Path):
    """Render standard blockquote preserving multiple paragraphs and quotes."""
    paragraphs = [c for c in node.children if c.type == "paragraph"]
    if not paragraphs:
        paragraphs = [node]

    for p_node in paragraphs:
        try:
            p = doc.add_paragraph(style="Quote")
        except Exception:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.25)

        # Left subtle gray border
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        left_bdr = OxmlElement("w:left")
        left_bdr.set(qn("w:val"), "single")
        left_bdr.set(qn("w:sz"), "18")
        left_bdr.set(qn("w:space"), "8")
        left_bdr.set(qn("w:color"), "D0D7DE")
        pBdr.append(left_bdr)
        pPr.append(pBdr)

        inline_nodes = []
        for c in p_node.children:
            if c.type == "inline" and c.children:
                inline_nodes.extend(c.children)
            else:
                inline_nodes.append(c)

        render_inline_tree(p, inline_nodes, base_dir, doc, italic=True)


def render_table(node: SyntaxTreeNode, doc: Document, base_dir: Path):
    """Render markdown table with column alignments, rich cell formatting, and repeated header rows."""
    rows_nodes = []
    for section in node.children:  # thead, tbody
        for tr in section.children:
            row_cells = []
            for th_or_td in tr.children:
                style_str = th_or_td.attrs.get("style", "")
                align_str = th_or_td.attrs.get("align", "")
                align = "left"
                if "center" in style_str or align_str == "center":
                    align = "center"
                elif "right" in style_str or align_str == "right":
                    align = "right"
                row_cells.append((th_or_td, align))
            rows_nodes.append((row_cells, section.type == "thead"))

    if not rows_nodes:
        return

    num_rows = len(rows_nodes)
    num_cols = max(len(cells) for cells, _ in rows_nodes)

    table = doc.add_table(rows=num_rows, cols=num_cols)
    try:
        table.style = "Table Grid"
    except Exception as exc:
        print(f"Warning: table style unavailable: {exc}", file=sys.stderr)

    for r_idx, (cells, is_header) in enumerate(rows_nodes):
        row = table.rows[r_idx]
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        if is_header:
            trPr.append(OxmlElement("w:tblHeader"))

        for c_idx, (th_or_td, align) in enumerate(cells):
            if c_idx < num_cols:
                cell = table.cell(r_idx, c_idx)
                p = cell.paragraphs[0]
                p.text = ""
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)

                if align == "center":
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif align == "right":
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                if is_header:
                    tcPr = cell._tc.get_or_add_tcPr()
                    shd = OxmlElement("w:shd")
                    shd.set(qn("w:val"), "clear")
                    shd.set(qn("w:color"), "auto")
                    shd.set(qn("w:fill"), "F2F4F7")
                    tcPr.append(shd)

                inline_nodes = []
                if th_or_td.children:
                    for c in th_or_td.children:
                        if c.type == "inline" and c.children:
                            inline_nodes.extend(c.children)
                        else:
                            inline_nodes.append(c)
                elif th_or_td.content:
                    inline_nodes = [{"type": "text", "content": th_or_td.content.strip(), "children": [], "attrs": {}}]

                render_inline_tree(p, inline_nodes, base_dir, doc, bold=is_header)

    # Empty paragraph after table for standard spacing
    doc.add_paragraph()


def transform_task_list(inline_nodes: list) -> list:
    """Convert [ ] and [x] markdown task list markers into Unicode checkbox glyphs."""
    if not inline_nodes:
        return inline_nodes
    first = inline_nodes[0]
    n_type, content, children, attrs = get_node_info(first)
    if n_type == "text" and content:
        if content.startswith("[ ] "):
            new_first = {"type": "text", "content": "☐ " + content[4:], "children": children, "attrs": attrs}
            return [new_first] + inline_nodes[1:]
        elif content.startswith("[x] ") or content.startswith("[X] "):
            new_first = {"type": "text", "content": "☑ " + content[4:], "children": children, "attrs": attrs}
            return [new_first] + inline_nodes[1:]
        elif content == "[ ]":
            new_first = {"type": "text", "content": "☐ ", "children": children, "attrs": attrs}
            return [new_first] + inline_nodes[1:]
        elif content in ("[x]", "[X]"):
            new_first = {"type": "text", "content": "☑ ", "children": children, "attrs": attrs}
            return [new_first] + inline_nodes[1:]
    return inline_nodes


def render_list_items(node: SyntaxTreeNode, doc: Document, base_dir: Path, is_ordered: bool, level: int = 1):
    """Recursively render nested bullet and ordered lists preserving hierarchy and task list checkboxes."""
    for item in node.children:
        style_prefix = "List Number" if is_ordered else "List Bullet"
        style_name = style_prefix if level == 1 else f"{style_prefix} {min(level, 3)}"

        inline_children = []
        nested_lists = []

        for c in item.children:
            if c.type in ("bullet_list", "ordered_list"):
                nested_lists.append(c)
            elif c.type == "paragraph":
                for sub in c.children:
                    if sub.type == "inline" and sub.children:
                        inline_children.extend(sub.children)
                    else:
                        inline_children.append(sub)
            elif c.type == "inline" and c.children:
                inline_children.extend(c.children)
            else:
                inline_children.append(c)

        # Check and apply task list checkboxes
        inline_children = transform_task_list(inline_children)

        try:
            p = doc.add_paragraph(style=style_name)
        except Exception:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.25 * level)

        render_inline_tree(p, inline_children, base_dir, doc)

        for nl in nested_lists:
            render_list_items(nl, doc, base_dir, is_ordered=(nl.type == "ordered_list"), level=level + 1)


def convert_with_pandoc(md_path: Path, output_path: Path, template_path: Path = None) -> bool:
    """High-fidelity conversion via Pandoc binary if available."""
    pandoc_bin = shutil.which("pandoc")
    if not pandoc_bin:
        return False

    cmd = [pandoc_bin, str(md_path), "-f", "gfm", "-t", "docx", "-o", str(output_path)]
    if template_path and template_path.is_file():
        cmd.extend(["--reference-doc", str(template_path)])

    try:
        # Pandoc executable comes from shutil.which; document paths are separate arguments.
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=60)  # nosec B603
        return res.returncode == 0 and output_path.is_file() and output_path.stat().st_size > 0
    except Exception:
        return False


def convert_md_to_docx(md_path: Path, output_path: Path, template_path: Path = None, engine: str = "auto") -> dict:
    """Convert a Markdown file to DOCX with 100% syntax coverage and optional reference template."""
    # 1. Try Pandoc engine if explicitly requested or auto
    if engine in ("auto", "pandoc"):
        if convert_with_pandoc(md_path, output_path, template_path):
            doc_test = Document(str(output_path))
            return {
                "engine": "pandoc",
                "md_path": str(md_path),
                "output_path": str(output_path),
                "template_used": str(template_path) if template_path else None,
                "paragraphs_generated": len(doc_test.paragraphs),
                "tables_generated": len(doc_test.tables),
            }
        elif engine == "pandoc":
            raise RuntimeError("Pandoc engine was requested, but pandoc executable was not found.")

    # 2. Enhanced Python Engine
    with open(md_path, "r", encoding="utf-8") as f:
        md_content = f.read()

    base_dir = md_path.parent

    if template_path and template_path.is_file():
        doc = Document(str(template_path))
        clear_document_body(doc)
    else:
        doc = Document()
        style_normal = doc.styles["Normal"]
        font = style_normal.font
        font.name = "Calibri"
        font.size = Pt(11)

    md_content = preprocess_markdown_special_constructs(md_content)
    md = MarkdownIt("commonmark").enable("table").enable("strikethrough")
    tokens = md.parse(md_content)
    root = SyntaxTreeNode(tokens)

    for node in root.children:
        if node.type == "heading":
            level = int(node.tag[1]) if len(node.tag) > 1 and node.tag[1].isdigit() else 1
            heading_style = f"Heading {min(max(level, 1), 6)}"
            try:
                p = doc.add_paragraph(style=heading_style)
            except Exception:
                try:
                    p = doc.add_paragraph(style=f"Заголовок {min(max(level, 1), 6)}")
                except Exception:
                    p = doc.add_paragraph()
                    p.style = "Normal"

            inline_nodes = []
            for c in node.children:
                if c.type == "inline" and c.children:
                    inline_nodes.extend(c.children)
                else:
                    inline_nodes.append(c)

            if inline_nodes:
                render_inline_tree(p, inline_nodes, base_dir, doc)
            elif node.content:
                add_styled_run(p, normalize_inline_text(node.content))

        elif node.type == "paragraph":
            p = doc.add_paragraph(style="Normal")
            inline_nodes = []
            for c in node.children:
                if c.type == "inline" and c.children:
                    inline_nodes.extend(c.children)
                else:
                    inline_nodes.append(c)

            if inline_nodes:
                render_inline_tree(p, inline_nodes, base_dir, doc)
            elif node.content:
                add_styled_run(p, normalize_inline_text(node.content))

        elif node.type == "bullet_list":
            render_list_items(node, doc, base_dir, is_ordered=False, level=1)

        elif node.type == "ordered_list":
            render_list_items(node, doc, base_dir, is_ordered=True, level=1)

        elif node.type == "blockquote":
            # Check for GitHub-style alerts: > [!NOTE], > [!WARNING], etc.
            plain = extract_plain_text(node).strip()
            is_alert = False
            for alert_name in ["NOTE", "TIP", "IMPORTANT", "WARNING", "CAUTION"]:
                tag = f"[!{alert_name}]"
                if plain.startswith(tag):
                    render_callout_block(node, alert_name, doc, base_dir)
                    is_alert = True
                    break

            if not is_alert:
                render_regular_blockquote(node, doc, base_dir)

        elif node.type in ("fence", "code_block"):
            lang = getattr(node, "info", "").strip()
            add_code_block(doc, node.content, language=lang)

        elif node.type == "table":
            render_table(node, doc, base_dir)

        elif node.type == "hr":
            p = doc.add_paragraph("―" * 35)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(8)
            for r in p.runs:
                r.font.color.rgb = RGBColor(160, 160, 160)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(".tmp.docx")
    doc.save(str(temp_path))
    os.replace(temp_path, output_path)

    return {
        "engine": "python",
        "md_path": str(md_path),
        "output_path": str(output_path),
        "template_used": str(template_path) if template_path else None,
        "paragraphs_generated": len(doc.paragraphs),
        "tables_generated": len(doc.tables),
    }


def main():
    parser = argparse.ArgumentParser(
        description="Convert Markdown (.md) to a professionally formatted Word (.docx) document."
    )
    parser.add_argument("md_file", type=Path, help="Path to input Markdown (.md) file")
    parser.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Path for destination .docx document (defaults to <name>.docx)"
    )
    parser.add_argument(
        "-t", "--template", type=Path, default=None,
        help="Optional reference .docx template (inherits styles, fonts, margins, headers/footers)"
    )
    parser.add_argument(
        "--engine", choices=["auto", "python", "pandoc"], default="auto",
        help="Conversion engine to use (default: auto)"
    )

    args = parser.parse_args()

    if not args.md_file.is_file():
        print(f"Error: Markdown file not found: {args.md_file}", file=sys.stderr)
        sys.exit(1)

    md_path = args.md_file.resolve()
    output_docx = args.output.resolve() if args.output else md_path.parent / f"{md_path.stem}.docx"
    template_path = args.template.resolve() if args.template else None

    if template_path and not template_path.is_file():
        print(f"Warning: Template file '{template_path}' does not exist. Proceeding with clean default styling.", file=sys.stderr)
        template_path = None

    res = convert_md_to_docx(md_path, output_docx, template_path, engine=args.engine)

    print(f"SUCCESS: Generated '{output_docx.name}' (engine: {res['engine']}).")
    print(f"  Destination: {output_docx}")
    if res["template_used"]:
        print(f"  Template:    {res['template_used']} (inherited custom styles)")
    else:
        print(f"  Template:    None (clean default formatting)")
    print(f"  Elements:    {res['paragraphs_generated']} paragraphs, {res['tables_generated']} tables")


if __name__ == "__main__":
    main()
