"""Input validation helpers that produce LLM-friendly error messages.

Every validation function raises ``ValueError`` with a message structured as:
    Line 1: what went wrong
    Line 2+: what the valid values are, or what to do instead

This gives LLMs clear, actionable feedback when a tool is called incorrectly.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import logging

from . import safety

log = logging.getLogger(__name__)

_CONTROL_CHAR_RE = re.compile(
    r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]"
)


def sanitize_text(text: str, field_name: str = "text") -> str:
    """Strip XML-illegal control characters from user input.

    OOXML forbids control chars 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F, 0x7F.
    Tab (0x09), newline (0x0A), carriage return (0x0D) are allowed.
    """
    if not isinstance(text, str):
        raise ValueError(
            f"{field_name} must be a string, got {type(text).__name__}.\n"
            f"Provide a text string value."
        )
    cleaned = _CONTROL_CHAR_RE.sub("", text)
    if len(cleaned) != len(text):
        removed = len(text) - len(cleaned)
        log.warning("Stripped %d control character(s) from %s", removed, field_name)
    return cleaned

_FORMAT_TOOLS = {
    "docx": {"create": "docx_create", "read": "docx_read", "append": "docx_append"},
    "xlsx": {"create": "xlsx_create", "read": "xlsx_read", "append": "xlsx_append_rows"},
    "pptx": {"create": "pptx_create", "read": "pptx_read", "append": "pptx_add_slide"},
}

_CELL_REF_RE = re.compile(r"^[A-Za-z]{1,3}\d{1,7}$")
_CELL_RANGE_RE = re.compile(r"^[A-Za-z]{1,3}\d{1,7}:[A-Za-z]{1,3}\d{1,7}$")

VALID_PAGE_SIZES = {
    "letter": (215.9, 279.4),
    "a4": (210.0, 297.0),
    "a3": (297.0, 420.0),
    "a5": (148.0, 210.0),
    "legal": (215.9, 355.6),
    "tabloid": (279.4, 431.8),
}


def validate_path(
    path: str,
    *,
    must_exist: bool = False,
    expected_ext: str | None = None,
    for_creation: bool = False,
    operation: str = "",
) -> Path:
    if not path or not path.strip():
        raise ValueError(
            "Path cannot be empty.\n"
            "Provide a valid file path, e.g. '/home/user/document.docx'"
        )

    p = Path(path).expanduser().resolve()
    ext = p.suffix.lstrip(".").lower()

    # Block access to system paths (/proc, /sys, /dev)
    safety.check_path_safety(p)

    if expected_ext:
        expected = expected_ext.lstrip(".").lower()
        if ext != expected:
            hint = ""
            if ext in _FORMAT_TOOLS:
                tools = _FORMAT_TOOLS[ext]
                if "read" in operation.lower():
                    hint = f"\nTo work with .{ext} files, use `{tools.get('read', 'N/A')}` instead."
                elif "create" in operation.lower() or for_creation:
                    hint = f"\nTo create a .{ext} file, use `{tools.get('create', 'N/A')}` instead."
                else:
                    hint = f"\nAvailable tools for .{ext}: {tools}"
            raise ValueError(
                f"Wrong file extension: expected .{expected}, got .{ext} for path '{p.name}'.\n"
                f"Either rename the file to end with .{expected}, or use the correct tool.{hint}"
            )

    if must_exist and not p.exists():
        suggestions = []
        if ext in _FORMAT_TOOLS:
            suggestions.append(
                f"Create it first with `{_FORMAT_TOOLS[ext].get('create', 'N/A')}`"
            )
        suggestions.append("Use `list_workspace` to see available documents in a directory")
        raise FileNotFoundError(
            f"File not found: {p}\n"
            + "\n".join(f"  - {s}" for s in suggestions)
        )

    if must_exist and p.exists() and not p.is_file():
        raise ValueError(f"Path is a directory, not a file: {p}")

    # OOM prevention: check file size and zip-bomb ratio before processing
    if must_exist and p.exists():
        safety.check_file_size(p)
        safety.check_zip_bomb(p)

    if for_creation:
        if not p.parent.exists():
            p.parent.mkdir(parents=True, exist_ok=True)

    return p


CSS_NAMED_COLORS: dict[str, str] = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000",
    "blue": "0000FF", "yellow": "FFFF00", "cyan": "00FFFF", "magenta": "FF00FF",
    "orange": "FFA500", "purple": "800080", "pink": "FFC0CB", "brown": "A52A2A",
    "gray": "808080", "grey": "808080", "silver": "C0C0C0", "gold": "FFD700",
    "navy": "000080", "teal": "008080", "olive": "808000", "maroon": "800000",
    "lime": "00FF00", "aqua": "00FFFF", "fuchsia": "FF00FF", "indigo": "4B0082",
    "coral": "FF7F50", "salmon": "FA8072", "crimson": "DC143C", "tomato": "FF6347",
    "chocolate": "D2691E", "sienna": "A0522D", "tan": "D2B48C", "wheat": "F5DEB3",
    "ivory": "FFFFF0", "beige": "F5F5DC", "linen": "FAF0E6", "snow": "FFFAFA",
    "darkblue": "00008B", "darkgreen": "006400", "darkred": "8B0000",
    "darkgray": "A9A9A9", "darkgrey": "A9A9A9", "darkviolet": "9400D3",
    "lightblue": "ADD8E6", "lightgreen": "90EE90", "lightgray": "D3D3D3",
    "lightgrey": "D3D3D3", "lightpink": "FFB6C1", "lightyellow": "FFFFE0",
    "steelblue": "4682B4", "royalblue": "4169E1", "skyblue": "87CEEB",
    "slategray": "708090", "slategrey": "708090", "dimgray": "696969",
    "forestgreen": "228B22", "seagreen": "2E8B57", "firebrick": "B22222",
    "midnightblue": "191970", "cornflowerblue": "6495ED",
}


def validate_color(color_hex: str) -> str:
    if not color_hex:
        raise ValueError(
            "Color cannot be empty.\n"
            "Provide a hex RGB color or named color.\n"
            "Examples: '#FF0000', 'red', 'navy', 'steelblue', '#FFF'."
        )
    stripped = color_hex.strip()
    name_lower = stripped.lower().lstrip("#")
    if name_lower in CSS_NAMED_COLORS:
        return CSS_NAMED_COLORS[name_lower]
    h = stripped.lstrip("#").upper()
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    if len(h) != 6:
        raise ValueError(
            f"Invalid color: '{color_hex}'. Expected 6-digit hex RGB or a named color.\n"
            f"Hex examples: '#FF0000', '#00FF00', '#0000FF', '#FFFFFF'.\n"
            f"Named colors: red, blue, green, navy, teal, coral, gold, steelblue, etc.\n"
            f"Full list: {', '.join(sorted(CSS_NAMED_COLORS.keys())[:20])}..."
        )
    try:
        int(h, 16)
    except ValueError:
        raise ValueError(
            f"Invalid hex digits in color: '{color_hex}'.\n"
            f"Use only hex characters 0-9 and A-F. Example: '#FF5733'.\n"
            f"Or use a named color: red, blue, green, navy, etc."
        )
    return h


def color_info() -> dict:
    """Return all supported named colors and usage guidance."""
    return {
        "named_colors": CSS_NAMED_COLORS,
        "count": len(CSS_NAMED_COLORS),
        "formats_accepted": ["#RRGGBB", "#RGB", "RRGGBB", "named (e.g. 'red', 'navy')"],
        "examples": {
            "hex_6": "#FF0000",
            "hex_3": "#F00",
            "bare_hex": "FF0000",
            "named": "red",
        },
    }


def validate_cell_ref(ref: str, param_name: str = "cell") -> str:
    if not ref:
        raise ValueError(
            f"{param_name} cannot be empty.\n"
            f"Provide an Excel-style cell reference, e.g. 'A1', 'B2', 'AA100'."
        )
    ref = ref.strip().upper()
    if not _CELL_REF_RE.match(ref):
        raise ValueError(
            f"Invalid cell reference: '{ref}'.\n"
            f"Use Excel-style format: column letter(s) + row number. "
            f"Examples: 'A1', 'B2', 'AA100', 'Z999'."
        )
    return ref


def validate_cell_range(range_str: str, param_name: str = "range") -> str:
    if not range_str:
        raise ValueError(
            f"{param_name} cannot be empty.\n"
            f"Provide an Excel-style range, e.g. 'A1:B10', 'C2:D20'."
        )
    range_str = range_str.strip().upper()
    if not _CELL_RANGE_RE.match(range_str):
        raise ValueError(
            f"Invalid cell range: '{range_str}'.\n"
            f"Use Excel-style format: 'START:END'. Examples: 'A1:B10', 'C2:D20', 'A1:Z100'."
        )
    return range_str


def validate_index(
    value: int,
    name: str,
    max_val: int,
    min_val: int = 0,
) -> int:
    if not isinstance(value, int):
        raise ValueError(
            f"{name} must be an integer, got {type(value).__name__}: {value!r}."
        )
    if value < min_val or value > max_val:
        raise ValueError(
            f"{name} = {value} is out of range [{min_val}, {max_val}].\n"
            f"The document has {max_val - min_val + 1} items (indexed {min_val} to {max_val})."
        )
    return value


def validate_sheet_name(wb: Any, sheet: str) -> str:
    if sheet not in wb.sheetnames:
        available = wb.sheetnames
        raise ValueError(
            f"Sheet '{sheet}' not found in workbook.\n"
            f"Available sheets: {available}\n"
            f"Sheet names are case-sensitive."
        )
    return sheet


def validate_paragraph_index(doc: Any, index: int) -> int:
    count = len(doc.paragraphs)
    if count == 0:
        raise ValueError(
            f"Document has no paragraphs. Add content first with `docx_append`."
        )
    return validate_index(index, "paragraph_index", max_val=count - 1)


def validate_slide_index(prs: Any, index: int) -> int:
    count = len(prs.slides)
    if count == 0:
        raise ValueError(
            f"Presentation has no slides. Add slides first with `pptx_add_slide`."
        )
    return validate_index(index, "slide_index", max_val=count - 1)


def validate_align(align: str) -> str:
    valid = {"left", "center", "right", "justify"}
    a = align.lower().strip()
    if a not in valid:
        raise ValueError(
            f"Invalid alignment: '{align}'.\n"
            f"Valid values: {sorted(valid)}"
        )
    return a


def validate_chart_type(chart_type: str) -> str:
    valid = {"bar", "line", "pie", "scatter", "area"}
    synonyms = {
        "column": "bar", "histogram": "bar",
        "doughnut": "pie", "donut": "pie",
        "xy": "scatter", "scatterplot": "scatter",
        "stacked_bar": "bar",
    }
    k = chart_type.lower().strip()
    k = synonyms.get(k, k)
    if k not in valid:
        raise ValueError(
            f"Unsupported chart type: '{chart_type}'.\n"
            f"Valid types: {sorted(valid)}\n"
            f"Synonyms: {synonyms}"
        )
    return k


def validate_page_size(size_name: str) -> tuple[float, float]:
    key = size_name.lower().strip()
    if key not in VALID_PAGE_SIZES:
        raise ValueError(
            f"Unknown page size: '{size_name}'.\n"
            f"Valid sizes: {sorted(VALID_PAGE_SIZES.keys())}\n"
            f"Or use docx_set_page_setup with custom width_mm/height_mm."
        )
    return VALID_PAGE_SIZES[key]


def validate_series_data(series: list[dict], chart_type: str) -> None:
    if not series:
        raise ValueError(
            "series cannot be empty. Provide at least one series dict.\n"
            'Example: [{"name": "Sales", "values": [10, 20, 30]}]'
        )
    ct = chart_type.lower()
    for i, s in enumerate(series):
        if not isinstance(s, dict):
            raise ValueError(
                f"series[{i}] must be a dict, got {type(s).__name__}.\n"
                f'Example: {{"name": "Sales", "values": [10, 20, 30]}}'
            )
        if ct == "scatter":
            if "x" not in s or "y" not in s:
                raise ValueError(
                    f"series[{i}] is missing 'x' or 'y' for scatter chart.\n"
                    f'Scatter series format: {{"name": "Series 1", "x": [1,2,3], "y": [4,5,6]}}'
                )
        else:
            if "values" not in s:
                raise ValueError(
                    f"series[{i}] is missing 'values' for {chart_type} chart.\n"
                    f'Series format: {{"name": "Sales", "values": [10, 20, 30]}}'
                )


def validate_choice(value, name: str, choices, *, lower: bool = True, allow_none: bool = True):
    """Validate that ``value`` is one of ``choices`` (case-insensitive by default).

    Returns the normalised value (lower-cased when ``lower``). ``None`` passes
    through when ``allow_none`` so optional parameters can be left unset.
    """
    if value is None:
        if allow_none:
            return None
        raise ValueError(f"{name} is required. Valid values: {sorted(choices)}")
    norm = value.lower().strip() if (lower and isinstance(value, str)) else value
    valid = {c.lower() if (lower and isinstance(c, str)) else c for c in choices}
    if norm not in valid:
        raise ValueError(
            f"Invalid {name}: {value!r}.\n"
            f"Valid values: {sorted(choices)}"
        )
    return norm


def validate_records(items, name: str, *, required=(), numeric=(), example: str = ""):
    """Validate a list-of-dicts argument (e.g. chart segments, card definitions).

    Ensures ``items`` is a non-empty list, every element is a dict, all
    ``required`` keys are present and non-empty, and every ``numeric`` key (when
    present) is coercible to a number. Raises AI-friendly ValueErrors that point
    at the offending index and show an example. Returns ``items`` unchanged.
    """
    hint = f"\nExample item: {example}" if example else ""
    if not isinstance(items, list) or not items:
        raise ValueError(
            f"{name} must be a non-empty list of dicts, got "
            f"{type(items).__name__ if not isinstance(items, list) else 'an empty list'}.{hint}"
        )
    for i, it in enumerate(items):
        if not isinstance(it, dict):
            raise ValueError(
                f"{name}[{i}] must be a dict, got {type(it).__name__}.{hint}"
            )
        for k in required:
            if k not in it or it[k] is None or it[k] == "":
                raise ValueError(
                    f"{name}[{i}] is missing required key '{k}' "
                    f"(keys present: {sorted(it.keys())}).{hint}"
                )
        for k in numeric:
            if k in it and it[k] is not None:
                try:
                    float(it[k])
                except (TypeError, ValueError):
                    raise ValueError(
                        f"{name}[{i}]['{k}'] must be a number, got {it[k]!r}.{hint}"
                    ) from None
    return items


def validate_slide_def(slide_def: dict) -> None:
    if not isinstance(slide_def, dict):
        raise ValueError(
            f"Slide definition must be a dict, got {type(slide_def).__name__}.\n"
            'Example: {"layout": "content", "title": "My Slide", "body": ["Bullet 1", "Bullet 2"]}\n'
            'Valid layouts: title, content, title_only, image, blank'
        )
    layout = slide_def.get("layout", "content")
    valid_layouts = {"title", "content", "title_only", "image", "blank"}
    if layout not in valid_layouts:
        raise ValueError(
            f"Invalid slide layout: '{layout}'.\n"
            f"Valid layouts: {sorted(valid_layouts)}\n"
            'Examples:\n'
            '  {"layout": "title", "title": "...", "subtitle": "..."}\n'
            '  {"layout": "content", "title": "...", "body": ["bullet 1", "bullet 2"]}\n'
            '  {"layout": "image", "title": "...", "image_path": "/path/to/image.png"}'
        )
    if layout == "image" and "image_path" in slide_def:
        img = Path(slide_def["image_path"]).expanduser().resolve()
        if not img.exists():
            raise FileNotFoundError(
                f"Slide image not found: {img}\n"
                f"Check the image path and try again."
            )


def validate_regex(pattern: str, max_length: int = 500) -> re.Pattern:
    if len(pattern) > max_length:
        raise ValueError(
            f"Regex pattern too long ({len(pattern)} chars, max {max_length}).\n"
            f"Use a shorter pattern, or set regex=False for literal search."
        )
    _dangerous = re.compile(
        r"(\(.+\+\)\+|\(\.\*\)\{|\(\[.*\]\+\)\+|"
        r"\(\.\+\)\+|\(\.\+\)\*|\(\.\*\)\+)"
    )
    if _dangerous.search(pattern):
        raise ValueError(
            f"Regex pattern rejected: nested quantifiers detected (risk of catastrophic "
            f"backtracking / ReDoS).\n"
            f"Simplify the pattern. Example bad: '(a+)+$'  Example good: 'a+'."
        )
    try:
        return re.compile(pattern)
    except re.error as e:
        raise ValueError(
            f"Invalid regex: {e}.\n"
            f"Check the pattern syntax. Set regex=False for literal string search."
        ) from e


def validate_bounded_int(
    value: int, name: str, *, min_val: int = 0, max_val: int = 100_000
) -> int:
    if not isinstance(value, int) or value < min_val or value > max_val:
        raise ValueError(
            f"{name}={value!r} out of range [{min_val}, {max_val}]."
        )
    return value


def format_for_path(path: str) -> str:
    return Path(path).suffix.lstrip(".").lower()


# ---------------------------------------------------------------------------
# Professional formatting validators
# ---------------------------------------------------------------------------

_LEGEND_POSITIONS = {
    "best", "upper right", "upper left", "lower left", "lower right",
    "right", "center left", "center right", "lower center", "upper center",
    "center",
}


def validate_line_spacing(val: float) -> float:
    if not isinstance(val, (int, float)) or val <= 0:
        raise ValueError(
            f"line_spacing must be a positive number, got {val!r}.\n"
            f"Common values: 1.0 (single), 1.15, 1.5 (one-and-a-half), 2.0 (double).\n"
            f"Values > 3 are treated as absolute point sizes."
        )
    return float(val)


def validate_legend_position(pos: str) -> str:
    p = pos.lower().strip()
    if p not in _LEGEND_POSITIONS:
        raise ValueError(
            f"Invalid legend position: '{pos}'.\n"
            f"Valid positions: {sorted(_LEGEND_POSITIONS)}"
        )
    return p


def validate_list_level(level: int) -> int:
    if not isinstance(level, int):
        level = int(level)
    return max(0, min(level, 5))


def validate_indent(val: float) -> float:
    if not isinstance(val, (int, float)):
        raise ValueError(f"Indent must be a number, got {type(val).__name__}.")
    if val < -144 or val > 720:
        raise ValueError(
            f"Indent value {val} out of range [-144, 720] points.\n"
            f"Common values: 36 (0.5 inch), 72 (1 inch), 144 (2 inches)."
        )
    return float(val)
