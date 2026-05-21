"""Chart insertion across docx / xlsx / pptx.

- **docx** -> render with matplotlib (Agg backend), save to a temp PNG,
  embed via ``python-docx.add_picture``. Result is a static image — NOT an
  editable chart. Documented in the tool docstring.
- **xlsx** -> native ``openpyxl`` charts. Editable inside the workbook.
- **pptx** -> native ``python-pptx`` charts. Editable inside the deck.

Force ``matplotlib.use("Agg")`` BEFORE any pyplot import so that chart
rendering works in headless / stdio-only contexts.
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Any

# Matplotlib must use a non-GUI backend.
import matplotlib

matplotlib.use("Agg")  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402


# --------------------------------------------------------------------------
# Supported chart kinds + synonyms
# --------------------------------------------------------------------------

_KINDS = ("bar", "line", "pie", "scatter", "area")
_SYNONYMS = {
    "column": "bar",
    "histogram": "bar",
    "doughnut": "pie",
    "donut": "pie",
    "xy": "scatter",
    "scatterplot": "scatter",
    "stacked_bar": "bar",
}


def _normalise_kind(kind: str) -> str:
    from .validation import validate_chart_type
    return validate_chart_type(kind)


def chart_kinds_info() -> dict:
    """Return supported chart kinds + synonyms + per-format support."""
    return {
        "kinds": list(_KINDS),
        "synonyms": dict(_SYNONYMS),
        "supported_in": {
            "docx": list(_KINDS),
            "xlsx": list(_KINDS),
            "pptx": ["bar", "line", "pie", "scatter"],  # area available but XL_CHART_TYPE varies
        },
    }


# --------------------------------------------------------------------------
# DOCX (matplotlib -> PNG -> add_picture)
# --------------------------------------------------------------------------

def docx_add_chart(
    path: str,
    chart_type: str,
    categories: list[Any],
    series: list[dict],
    *,
    title: str | None = None,
    width_inches: float = 6.0,
    height_inches: float = 4.0,
    paragraph_index: int | None = None,
) -> str:
    """Render a chart with matplotlib and embed it as a picture in a .docx.

    ``series`` is a list of ``{"name": str, "values": [numbers]}`` for
    bar/line/area/pie, or ``{"name": str, "x": [numbers], "y": [numbers]}``
    for scatter.

    Note: the embedded result is a static image — the chart cannot be edited
    inside Word/OnlyOffice. Use ``xlsx_add_chart`` / ``pptx_add_chart`` for
    editable charts.
    """
    from docx import Document
    from docx.shared import Inches

    from .validation import validate_series_data

    kind = _normalise_kind(chart_type)
    validate_series_data(series, kind)
    out = Path(path).expanduser().resolve()
    if not out.exists():
        raise FileNotFoundError(out)

    fig, ax = plt.subplots(figsize=(width_inches, height_inches), dpi=120)

    if kind == "pie":
        # Pie uses only the first series' values + the categories as labels.
        values = list(series[0].get("values", []))
        labels = [str(c) for c in categories[: len(values)]]
        ax.pie(values, labels=labels, autopct="%1.1f%%", startangle=90)
        ax.axis("equal")
    elif kind == "scatter":
        for s in series:
            ax.scatter(s.get("x", []), s.get("y", []), label=s.get("name"))
        ax.legend(loc="best")
    elif kind == "line":
        for s in series:
            ax.plot(categories, s.get("values", []), label=s.get("name"), marker="o")
        ax.legend(loc="best")
    elif kind == "area":
        bottom = [0] * len(categories)
        for s in series:
            values = list(s.get("values", []))
            ax.fill_between(categories, bottom, [b + v for b, v in zip(bottom, values)],
                            label=s.get("name"), alpha=0.6)
            bottom = [b + v for b, v in zip(bottom, values)]
        ax.legend(loc="best")
    else:  # bar
        n = len(series)
        width = 0.8 / max(1, n)
        x_positions = list(range(len(categories)))
        for i, s in enumerate(series):
            offsets = [x + (i - (n - 1) / 2) * width for x in x_positions]
            ax.bar(offsets, s.get("values", []), width=width, label=s.get("name"))
        ax.set_xticks(x_positions)
        ax.set_xticklabels([str(c) for c in categories])
        ax.legend(loc="best")

    if title:
        ax.set_title(title)
    fig.tight_layout()

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp_path = tmp.name
        fig.savefig(tmp_path, bbox_inches="tight")
        plt.close(fig)

        doc = Document(str(out))
        if paragraph_index is not None and 0 <= paragraph_index < len(doc.paragraphs):
            target = doc.paragraphs[paragraph_index]
            doc.add_picture(tmp_path, width=Inches(width_inches))
            pic_para = doc.element.body[-1]
            doc.element.body.remove(pic_para)
            target._p.addnext(pic_para)
        else:
            doc.add_picture(tmp_path, width=Inches(width_inches))
        doc.save(str(out))
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    return str(out)


# --------------------------------------------------------------------------
# XLSX (native openpyxl)
# --------------------------------------------------------------------------

def xlsx_add_chart(
    path: str,
    sheet: str,
    chart_type: str,
    data_range: str,
    *,
    categories_range: str | None = None,
    anchor_cell: str = "E2",
    title: str | None = None,
) -> str:
    """Add a native chart to an existing .xlsx file.

    ``data_range`` and ``categories_range`` are A1-style refs *within the
    sheet*: e.g. ``"B1:B10"`` (first row is treated as series title). If
    ``categories_range`` is omitted, the first column of ``data_range`` is
    used as categories.
    """
    from openpyxl import load_workbook
    from openpyxl.chart import (
        AreaChart,
        BarChart,
        LineChart,
        PieChart,
        Reference,
        ScatterChart,
    )

    from .validation import validate_cell_range, validate_sheet_name

    kind = _normalise_kind(chart_type)
    validate_cell_range(data_range, "data_range")
    if categories_range:
        validate_cell_range(categories_range, "categories_range")
    out = Path(path).expanduser().resolve()
    if not out.exists():
        raise FileNotFoundError(out)

    wb = load_workbook(str(out))
    validate_sheet_name(wb, sheet)
    ws = wb[sheet]

    # Pre-extend the sheet by writing None at the anchor cell to avoid Excel
    # rejecting under-sized sheets.
    if ws[anchor_cell].value is None:
        ws[anchor_cell] = None

    chart_cls = {
        "bar": BarChart,
        "line": LineChart,
        "pie": PieChart,
        "scatter": ScatterChart,
        "area": AreaChart,
    }[kind]
    chart = chart_cls()
    if title:
        chart.title = title

    # Reference accepts a range_string like "Sheet1!B1:C10".
    chart.add_data(
        Reference(ws, range_string=f"{sheet}!{data_range}"),
        titles_from_data=True,
    )
    if categories_range:
        chart.set_categories(Reference(ws, range_string=f"{sheet}!{categories_range}"))

    ws.add_chart(chart, anchor_cell)
    wb.save(str(out))
    return str(out)


# --------------------------------------------------------------------------
# PPTX (native python-pptx)
# --------------------------------------------------------------------------

def pptx_add_chart(
    path: str,
    slide_index: int,
    chart_type: str,
    categories: list[Any],
    series: list[dict],
    *,
    left_inches: float = 1.0,
    top_inches: float = 2.0,
    width_inches: float = 8.0,
    height_inches: float = 5.0,
    title: str | None = None,
) -> str:
    """Add a native chart to an existing slide in a .pptx file.

    For bar / line / pie / area: ``series`` is a list of ``{"name", "values"}``.
    For scatter: ``series`` is a list of ``{"name", "x", "y"}``.
    """
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData, XyChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches

    from .validation import validate_series_data

    kind = _normalise_kind(chart_type)
    validate_series_data(series, kind)
    kind_map = {
        "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "line": XL_CHART_TYPE.LINE,
        "pie": XL_CHART_TYPE.PIE,
        "scatter": XL_CHART_TYPE.XY_SCATTER,
        "area": XL_CHART_TYPE.AREA,
    }
    if kind not in kind_map:
        raise ValueError(f"Chart kind '{kind}' not supported in pptx")

    out = Path(path).expanduser().resolve()
    if not out.exists():
        raise FileNotFoundError(out)

    prs = Presentation(str(out))
    if slide_index < 0 or slide_index >= len(prs.slides):
        raise ValueError(f"slide_index {slide_index} out of range [0, {len(prs.slides) - 1}]")
    slide = prs.slides[slide_index]

    if kind == "scatter":
        chart_data = XyChartData()
        for s in series:
            ser = chart_data.add_series(s.get("name", "Series"))
            xs = s.get("x", [])
            ys = s.get("y", [])
            for x_val, y_val in zip(xs, ys):
                ser.add_data_point(x_val, y_val)
    else:
        chart_data = CategoryChartData()
        chart_data.categories = [str(c) for c in categories]
        for s in series:
            chart_data.add_series(s.get("name", "Series"), s.get("values", []))

    chart_shape = slide.shapes.add_chart(
        kind_map[kind],
        Inches(left_inches),
        Inches(top_inches),
        Inches(width_inches),
        Inches(height_inches),
        chart_data,
    )
    if title:
        chart_shape.chart.has_title = True
        chart_shape.chart.chart_title.text_frame.text = title

    prs.save(str(out))
    return str(out)
