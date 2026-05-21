"""Excel workbook operations using openpyxl."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


def _sanitize_sheet_name(name: str) -> str:
    # Excel sheet names: max 31 chars, no [ ] : * ? / \
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
    out = Path(path).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    # Remove the default blank sheet — we'll add ours.
    default_sheet = wb.active
    wb.remove(default_sheet)

    if not sheets:
        sheets = {"Sheet1": []}

    for name, rows in sheets.items():
        ws = wb.create_sheet(title=_sanitize_sheet_name(name))
        for row in rows:
            ws.append(list(row))
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
    in_ = Path(path).expanduser().resolve()
    if not in_.exists():
        raise FileNotFoundError(in_)
    wb = load_workbook(str(in_), data_only=True)
    if sheet is not None:
        if sheet not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet}' not in workbook. Available: {wb.sheetnames}")
        ws = wb[sheet]
        return {"sheet": sheet, "rows": [list(row) for row in ws.iter_rows(values_only=True)]}
    return {
        name: [list(row) for row in wb[name].iter_rows(values_only=True)]
        for name in wb.sheetnames
    }


def append_rows(path: str, sheet: str, rows: list[list[Any]]) -> str:
    """Append rows to an existing sheet. Creates the sheet if missing."""
    in_ = Path(path).expanduser().resolve()
    if not in_.exists():
        raise FileNotFoundError(in_)
    wb = load_workbook(str(in_))
    target = _sanitize_sheet_name(sheet)
    if target not in wb.sheetnames:
        ws = wb.create_sheet(title=target)
    else:
        ws = wb[target]
    for row in rows:
        ws.append(list(row))
    wb.save(str(in_))
    return str(in_)


def set_cell(path: str, sheet: str, cell: str, value: Any) -> str:
    """Set a single cell (A1, B2, etc.) in an existing workbook."""
    in_ = Path(path).expanduser().resolve()
    if not in_.exists():
        raise FileNotFoundError(in_)
    wb = load_workbook(str(in_))
    if sheet not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet}' not found. Available: {wb.sheetnames}")
    wb[sheet][cell] = value
    wb.save(str(in_))
    return str(in_)


def list_sheets(path: str) -> list[str]:
    """Return the sheet names in a workbook."""
    in_ = Path(path).expanduser().resolve()
    if not in_.exists():
        raise FileNotFoundError(in_)
    return load_workbook(str(in_), read_only=True).sheetnames
