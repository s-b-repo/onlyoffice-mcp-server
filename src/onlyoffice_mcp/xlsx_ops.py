"""Excel workbook operations using openpyxl."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

from .validation import validate_path, validate_cell_ref, validate_sheet_name, sanitize_text


def _sanitize_sheet_name(name: str) -> str:
    # Excel sheet names: max 31 chars, no [ ] : * ? / \
    name = sanitize_text(name, "sheet name")
    forbidden = '[]:*?/\\'
    cleaned = "".join("_" if c in forbidden else c for c in name)
    return cleaned[:31] or "Sheet1"


def create(
    path: str,
    sheets: dict[str, list[list[Any]]],
    *,
    header_bold: bool = True,
) -> str:
    """Create an .xlsx workbook at `path`.

    `sheets` maps sheet name -> rows (list of row lists). If empty, a single
    blank sheet named 'Sheet1' is created. The first row of each sheet is
    bolded when `header_bold` is True.
    """
    out = validate_path(path, expected_ext="xlsx", for_creation=True, operation="create")
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    if not sheets:
        sheets = {"Sheet1": []}

    for name, rows in sheets.items():
        ws = wb.create_sheet(title=_sanitize_sheet_name(name))
        for row in rows:
            ws.append([sanitize_text(str(c), "cell") if isinstance(c, str) else c for c in row])
        if header_bold and rows:
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="left")

    wb.save(str(out))
    return str(out)


def read(path: str, sheet: str | None = None) -> Any:
    """Read an .xlsx workbook.

    If `sheet` is provided, returns {'sheet': <name>, 'rows': [[...]]}.
    Otherwise returns {sheet_name: [[...]]} for every sheet.
    """
    in_ = validate_path(path, must_exist=True, expected_ext="xlsx", operation="read")
    wb = load_workbook(str(in_), data_only=True)
    if sheet is not None:
        validate_sheet_name(wb, sheet)
        ws = wb[sheet]
        return {"sheet": sheet, "rows": [list(row) for row in ws.iter_rows(values_only=True)]}
    return {
        name: [list(row) for row in wb[name].iter_rows(values_only=True)]
        for name in wb.sheetnames
    }


def append_rows(path: str, sheet: str, rows: list[list[Any]]) -> str:
    """Append rows to an existing sheet. Creates the sheet if missing."""
    in_ = validate_path(path, must_exist=True, expected_ext="xlsx", operation="append_rows")
    wb = load_workbook(str(in_))
    target = _sanitize_sheet_name(sheet)
    if target not in wb.sheetnames:
        ws = wb.create_sheet(title=target)
    else:
        ws = wb[target]
    for row in rows:
        ws.append([sanitize_text(str(c), "cell") if isinstance(c, str) else c for c in row])
    wb.save(str(in_))
    return str(in_)


def set_cell(path: str, sheet: str, cell: str, value: Any) -> str:
    """Set a single cell (A1, B2, etc.) in an existing workbook."""
    in_ = validate_path(path, must_exist=True, expected_ext="xlsx", operation="set_cell")
    cell = validate_cell_ref(cell)
    wb = load_workbook(str(in_))
    validate_sheet_name(wb, sheet)
    if isinstance(value, str):
        value = sanitize_text(value, "cell value")
    wb[sheet][cell] = value
    wb.save(str(in_))
    return str(in_)


def list_sheets(path: str) -> list[str]:
    """Return the sheet names in a workbook."""
    in_ = validate_path(path, must_exist=True, expected_ext="xlsx", operation="list_sheets")
    return load_workbook(str(in_), read_only=True).sheetnames


def delete_sheet(path: str, sheet: str) -> str:
    """Delete a sheet from an existing workbook."""
    in_ = validate_path(path, must_exist=True, expected_ext="xlsx", operation="delete_sheet")
    wb = load_workbook(str(in_))
    validate_sheet_name(wb, sheet)
    if len(wb.sheetnames) <= 1:
        raise ValueError(
            "Cannot delete the last sheet in a workbook.\n"
            "A workbook must have at least one sheet."
        )
    del wb[sheet]
    wb.save(str(in_))
    return str(in_)


def rename_sheet(path: str, old_name: str, new_name: str) -> str:
    """Rename a sheet in an existing workbook."""
    in_ = validate_path(path, must_exist=True, expected_ext="xlsx", operation="rename_sheet")
    wb = load_workbook(str(in_))
    validate_sheet_name(wb, old_name)
    if new_name in wb.sheetnames:
        raise ValueError(
            f"Sheet '{new_name}' already exists.\n"
            f"Choose a different name. Current sheets: {wb.sheetnames}"
        )
    wb[old_name].title = _sanitize_sheet_name(new_name)
    wb.save(str(in_))
    return str(in_)


def delete_rows(path: str, sheet: str, start_row: int, count: int = 1) -> str:
    """Delete rows from a sheet. ``start_row`` is 1-based."""
    in_ = validate_path(path, must_exist=True, expected_ext="xlsx", operation="delete_rows")
    wb = load_workbook(str(in_))
    validate_sheet_name(wb, sheet)
    ws = wb[sheet]
    if start_row < 1 or start_row > (ws.max_row or 1):
        raise ValueError(
            f"start_row={start_row} is out of range [1, {ws.max_row or 1}].\n"
            f"Row numbers are 1-based."
        )
    if count < 1:
        raise ValueError("count must be at least 1.")
    ws.delete_rows(start_row, count)
    wb.save(str(in_))
    return str(in_)


def merge_cells(path: str, sheet: str, range_str: str) -> str:
    """Merge cells in a range (e.g. 'A1:C1')."""
    from .validation import validate_cell_range
    in_ = validate_path(path, must_exist=True, expected_ext="xlsx", operation="merge_cells")
    range_str = validate_cell_range(range_str, "range")
    wb = load_workbook(str(in_))
    validate_sheet_name(wb, sheet)
    wb[sheet].merge_cells(range_str)
    wb.save(str(in_))
    return str(in_)


def insert_rows(path: str, sheet: str, row: int, count: int = 1) -> str:
    """Insert blank rows before ``row`` (1-based)."""
    in_ = validate_path(path, must_exist=True, expected_ext="xlsx", operation="insert_rows")
    wb = load_workbook(str(in_))
    validate_sheet_name(wb, sheet)
    ws = wb[sheet]
    if row < 1:
        raise ValueError(
            f"row={row} must be >= 1.\n"
            f"Row numbers are 1-based."
        )
    if count < 1:
        raise ValueError("count must be at least 1.")
    ws.insert_rows(row, count)
    wb.save(str(in_))
    return str(in_)


def set_column_width(path: str, sheet: str, column: str, width: float) -> str:
    """Set the width of a column (e.g. column='A', width=20)."""
    in_ = validate_path(path, must_exist=True, expected_ext="xlsx", operation="set_column_width")
    wb = load_workbook(str(in_))
    validate_sheet_name(wb, sheet)
    col = column.strip().upper()
    if not col.isalpha() or len(col) > 3:
        raise ValueError(
            f"Invalid column letter: '{column}'.\n"
            f"Use a column letter like 'A', 'B', 'AA'. Not a cell reference."
        )
    if width <= 0 or width > 255:
        raise ValueError(
            f"width={width} is out of range (0, 255].\n"
            f"Typical column widths: 8 (default), 12, 20, 30."
        )
    wb[sheet].column_dimensions[col].width = width
    wb.save(str(in_))
    return str(in_)


def freeze_panes(path: str, sheet: str, cell: str) -> str:
    """Freeze rows/columns above and to the left of ``cell``.

    For example, freeze_panes(path, sheet, 'A2') freezes the first row.
    Use 'B2' to freeze the first row and first column.
    """
    in_ = validate_path(path, must_exist=True, expected_ext="xlsx", operation="freeze_panes")
    cell = validate_cell_ref(cell)
    wb = load_workbook(str(in_))
    validate_sheet_name(wb, sheet)
    wb[sheet].freeze_panes = cell
    wb.save(str(in_))
    return str(in_)
