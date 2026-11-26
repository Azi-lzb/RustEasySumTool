"""Small, self-contained spreadsheet style helpers shared by demos and reports.

The professional palette and sizing conventions follow the XLSX design tokens.
There are deliberately no imports from a developer machine or a ZCode plugin.
"""

from __future__ import annotations

from copy import copy
from decimal import Decimal
import math
import os
from pathlib import Path
import unicodedata
from typing import TYPE_CHECKING

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

if TYPE_CHECKING:
    from openpyxl.cell.cell import Cell
    from openpyxl.worksheet.worksheet import Worksheet


PRIMARY = "1B2A4A"
PRIMARY_LIGHT = "D6E4F0"
SECONDARY = PRIMARY_LIGHT
ACCENT_POSITIVE = "1B7D46"
ACCENT_NEGATIVE = "C0392B"
ACCENT_WARNING = "D4820A"
NEUTRAL_900 = "37352F"
NEUTRAL_600 = "8C8A84"
NEUTRAL_200 = "E9E9E8"
NEUTRAL_100 = "F7F7F5"
NEUTRAL_50 = "FAFAF9"
NEUTRAL_0 = "FFFFFF"
HEADER_TEXT = NEUTRAL_0


def _local_font() -> str:
    fonts = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
    candidates = (
        (fonts / "msyh.ttc", "微软雅黑"),
        (fonts / "Deng.ttf", "等线"),
        (Path("/System/Library/Fonts/PingFang.ttc"), "PingFang SC"),
        (Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"), "Noto Sans CJK SC"),
        (Path("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"), "WenQuanYi Micro Hei"),
    )
    return next((name for path, name in candidates if path.is_file()), "Noto Sans CJK SC")


FONT_NAME = _local_font()
HEADER_BOLD = FONT_NAME not in {"Noto Sans SC", "Noto Sans CJK SC", "WenQuanYi Micro Hei"}
FORMATS = {
    "integer": "#,##0",
    "decimal_1": "#,##0.0",
    "decimal_2": "#,##0.00",
    "percentage": "0.0%",
    "currency_cny": "¥#,##0.00",
    "currency_usd": "$#,##0.00",
    "date": "YYYY-MM-DD",
}
ROW_HEIGHTS = {"margin": 15, "title": 32, "spacer": 8, "header": 28, "data": 22, "total": 26}
COLUMN_WIDTHS = {"margin": 3, "name_cn": 16, "number": 14, "percentage": 12, "description": 32}


def font_body() -> Font:
    return Font(name=FONT_NAME, size=11, color=NEUTRAL_900)


def font_caption() -> Font:
    return Font(name=FONT_NAME, size=9, color=NEUTRAL_600)


def align_number() -> Alignment:
    return Alignment(horizontal="right", vertical="center", wrap_text=True)


def align_text() -> Alignment:
    return Alignment(horizontal="left", vertical="center", wrap_text=True)


def setup_sheet(ws: Worksheet, title: str | None = None, last_col: int | None = None) -> None:
    """Use a B2 canvas, native Chinese font, frozen labels and A4 print defaults.

    Callers can set their final print_area after writing all rows. Existing data
    dimensions are not resized here, apart from the standard A/1/2/3 margins.
    """
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = COLUMN_WIDTHS["margin"]
    for row, role in ((1, "margin"), (2, "title"), (3, "spacer")):
        ws.row_dimensions[row].height = ROW_HEIGHTS[role]
    ws.freeze_panes = "C5"
    ws.print_title_rows = "4:4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.4, bottom=0.4, header=0.15, footer=0.15)
    ws.oddFooter.center.text = "第 &P 页 / 共 &N 页"
    ws.oddFooter.center.size = 9
    ws.oddFooter.center.font = FONT_NAME
    if title is not None:
        if last_col is not None and last_col > 2:
            ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=last_col)
        cell = ws.cell(row=2, column=2, value=title)
        cell.font = Font(name=FONT_NAME, size=16, bold=HEADER_BOLD, color=PRIMARY)
        cell.alignment = Alignment(horizontal="left", vertical="center")


def style_header_row(ws: Worksheet, row_num: int, col_start: int, col_end: int) -> None:
    for col in range(col_start, col_end + 1):
        cell = ws.cell(row_num, col)
        cell.fill = PatternFill("solid", fgColor=PRIMARY)
        cell.font = Font(name=FONT_NAME, size=11, bold=HEADER_BOLD, color=HEADER_TEXT)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin", color=NEUTRAL_200))
    ws.row_dimensions[row_num].height = ROW_HEIGHTS["header"]


def style_data_row(
    ws: Worksheet,
    row_num: int,
    col_start: int,
    col_end: int,
    row_index: int | None = None,
) -> None:
    index = row_num - 5 if row_index is None else row_index
    fill = PatternFill("solid", fgColor=NEUTRAL_0 if index % 2 == 0 else NEUTRAL_100)
    for col in range(col_start, col_end + 1):
        cell = ws.cell(row_num, col)
        cell.fill = fill
        cell.font = font_body()
        alignment = copy(cell.alignment)
        alignment.horizontal = alignment.horizontal or "left"
        alignment.vertical = alignment.vertical or "center"
        alignment.wrap_text = True
        cell.alignment = alignment
        cell.border = Border()
    ws.row_dimensions[row_num].height = ROW_HEIGHTS["data"]


def style_total_row(ws: Worksheet, row_num: int, col_start: int, col_end: int) -> None:
    for col in range(col_start, col_end + 1):
        cell = ws.cell(row_num, col)
        cell.fill = PatternFill("solid", fgColor=SECONDARY)
        cell.font = Font(name=FONT_NAME, size=11, bold=HEADER_BOLD, color=PRIMARY)
        cell.border = Border(top=Side(style="medium", color=NEUTRAL_200))
        alignment = copy(cell.alignment)
        alignment.horizontal = alignment.horizontal or "left"
        alignment.vertical = alignment.vertical or "center"
        alignment.wrap_text = True
        cell.alignment = alignment
    ws.row_dimensions[row_num].height = ROW_HEIGHTS["total"]


def _display_width(text: object) -> float:
    return sum(1.7 if unicodedata.east_asian_width(char) in {"W", "F"} else 1 for char in str(text))


def _display_text(cell: Cell) -> str:
    value = cell.value
    if value is None:
        return ""
    if cell.data_type == "f":
        # Excel displays the result, not a potentially very long formula.
        return "0.00%" if "%" in cell.number_format else "0,000,000.00"
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        if "%" in cell.number_format:
            return f"{value * 100:.2f}%"
        if "#,##" in cell.number_format:
            return f"{value:,.2f}"
    return str(value)


def _header_row(ws: Worksheet, requested: int | None) -> int | None:
    if requested is not None:
        return requested
    if ws.cell(4, 2).value is not None:
        return 4
    return next((row for row in range(3, ws.max_row + 1) if ws.cell(row, 2).value is not None), None)


def auto_fit_columns(
    ws: Worksheet,
    min_width: float = 8,
    max_width: float = 28,
    header_row: int | None = None,
    data_start_row: int | None = None,
) -> None:
    """Estimate widths from body data; merged titles/captions do not widen B."""
    if not 0 < min_width <= max_width:
        raise ValueError("列宽范围须满足 0 < min_width <= max_width")
    header = _header_row(ws, header_row)
    if header is None:
        return
    start = header + 1 if data_start_row is None else data_start_row
    merged = tuple(ws.merged_cells.ranges)
    for col in range(2, ws.max_column + 1):
        widths = []
        for row in range(start, ws.max_row + 1):
            cell = ws.cell(row, col)
            if any(cell.coordinate in area and area.max_col > area.min_col for area in merged):
                continue
            widths.extend(_display_width(line) for line in _display_text(cell).splitlines())
        width = min(max_width, max(min_width, max(widths, default=0) + 2))
        ws.column_dimensions[get_column_letter(col)].width = width
        alignment = copy(ws.cell(header, col).alignment)
        alignment.wrap_text = True
        ws.cell(header, col).alignment = alignment


def auto_fit_row_heights(
    ws: Worksheet,
    header_row: int | None = None,
    data_start_row: int | None = None,
    data_end_row: int | None = None,
    base_height: float = 22,
    line_height: float = 16,
    font_size: float = 11,
    min_height: float = 22,
) -> None:
    """Grow wrapped rows after final column sizing, including merged captions."""
    header = _header_row(ws, header_row)
    if header is None:
        return
    start = header + 1 if data_start_row is None else data_start_row
    end = ws.max_row if data_end_row is None else data_end_row
    merged = tuple(ws.merged_cells.ranges)
    rows = dict.fromkeys((header, *range(start, end + 1)))
    for row in rows:
        required = ROW_HEIGHTS["header"] if row == header else min_height
        for col in range(2, ws.max_column + 1):
            cell = ws.cell(row, col)
            if cell.value is None or not (cell.alignment.wrap_text or "\n" in str(cell.value)):
                continue
            area = next((item for item in merged if cell.coordinate in item), None)
            if area is not None and (row, col) != (area.min_row, area.min_col):
                continue
            cols = range(area.min_col, area.max_col + 1) if area is not None else (col,)
            width = sum(ws.column_dimensions[get_column_letter(item)].width or 11 for item in cols)
            available = max(1, width * 0.9 * 11 / max(1, font_size))
            lines = sum(max(1, math.ceil(_display_width(line) / available)) for line in _display_text(cell).split("\n"))
            required = max(required, base_height + (lines - 1) * line_height)
        ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 0, required)


__all__ = [
    "FONT_NAME", "FORMATS", "PRIMARY", "PRIMARY_LIGHT", "SECONDARY",
    "ACCENT_POSITIVE", "ACCENT_NEGATIVE", "ACCENT_WARNING",
    "NEUTRAL_900", "NEUTRAL_600", "NEUTRAL_200", "NEUTRAL_100", "NEUTRAL_50", "NEUTRAL_0",
    "setup_sheet", "style_header_row", "style_data_row", "style_total_row",
    "auto_fit_columns", "auto_fit_row_heights", "align_number", "align_text", "font_caption", "font_body",
]
