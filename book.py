"""Build the workbook: a grid of formula cells that repaints when `n` changes.

Three sheets:

    Grid   one cell per pixel. Each holds =INDEX(Data!<its own row>, k)
    Data   one row per grid cell, one column per DISTINCT frame
    Map    one row: output frame n -> the Data column that draws it

Colour comes from conditional formatting - one rule per palette entry, keyed on
the cell's own value - because a formula cannot set a fill. That indirection is
the whole trick: changing `n` changes 8,160 numbers, and Excel repaints them.
"""

from __future__ import annotations

import numpy as np
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.formatting.rule import Rule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

# Excel measures column width in characters of the Normal font's widest digit
# and row height in points, which do not share a unit. Getting SQUARE cells
# means pinning both to one target, and the conversion is NOT the documented
# ECMA-376 one - that formula (`Truncate((px-5)/MDW*100+0.5)/100`, MDW=7) is a
# Windows 96-dpi convention. Measured against Excel 16.112 on macOS by writing
# known widths and asking Excel for the rendered size of the cell:
#
#   width attr  1.0  1.5  2.0  3.0  4.0  6.0  8.0  12.0
#   rendered pt 6    9    12   18   24   36   48   72
#
# which is exactly `rendered_points = 6 * width`, straight through the origin,
# while row height in points renders 1:1. Trusting the ECMA formula instead
# produced 9x12pt cells - a picture stretched 33% vertically.
PT_PER_WIDTH = 6.0
MAX_ROWS, MAX_COLS = 1_048_576, 16_384

GRID_TOP = 3          # first sheet row of the picture; rows 1-2 are the controls


def cell_geometry(pt: float) -> tuple[float, float]:
    """(column width attribute, row height in points) for a square cell."""
    return pt / PT_PER_WIDTH, float(pt)


def geometry_px(width_attr: float, height_pt: float) -> tuple[int, int]:
    """Invert cell_geometry - what Excel actually renders, in points."""
    return round(width_attr * PT_PER_WIDTH), round(height_pt)


def check_limits(w: int, h: int, ndistinct: int) -> None:
    """Assert the sheet actually fits, rather than assuming it does."""
    rows, cols = w * h, ndistinct
    if rows > MAX_ROWS:
        raise SystemExit(f"Data needs {rows:,} rows, Excel allows {MAX_ROWS:,}")
    if cols > MAX_COLS:
        raise SystemExit(f"Data needs {cols:,} columns, Excel allows {MAX_COLS:,}")
    if h + GRID_TOP - 1 > MAX_ROWS or w > MAX_COLS:
        raise SystemExit(f"grid {w}x{h} does not fit the sheet")


def write_workbook(path: str, frames: list[np.ndarray], order: list[int],
                   palette_bgr: np.ndarray, cell_px: int = 16,
                   title: str = "", width_scale: float = 1.0) -> dict:
    """Write the .xlsx. `frames` are distinct index images, `order` maps n->frame.

    `width_scale` corrects the column width so the cells render square - it is
    measured from a screenshot rather than derived, see capture.fit_cells.
    """
    h, w = frames[0].shape
    check_limits(w, h, len(frames))
    cw, rh = cell_geometry(cell_px)
    cw *= width_scale
    last_col = get_column_letter(w)
    data_last = get_column_letter(len(frames))
    map_last = get_column_letter(len(order))

    wb = Workbook(write_only=True)
    wb.calculation.fullCalcOnLoad = True
    grid = wb.create_sheet("Grid")
    data = wb.create_sheet("Data")
    fmap = wb.create_sheet("Map")

    # --- Grid -------------------------------------------------------------
    grid.sheet_format.defaultColWidth = cw
    grid.sheet_format.defaultRowHeight = rh
    grid.sheet_format.customHeight = True
    for c in range(1, w + 1):
        grid.column_dimensions[get_column_letter(c)].width = cw
    for r in range(GRID_TOP, GRID_TOP + h):
        grid.row_dimensions[r].height = rh

    # The two control cells sit above the picture. Their columns are only a few
    # pixels wide, so a bare number would render as "#" - they are merged into a
    # readable block instead.
    span = max(6, round(220 / cell_px))
    for r in (1, 2):
        grid.merged_cells.ranges.add(f"A{r}:{get_column_letter(span)}{r}")
        grid.row_dimensions[r].height = 22

    def control(value, note):
        c = WriteOnlyCell(grid, value=value)
        c.font = Font(bold=True, size=14)
        c.alignment = Alignment(horizontal="center", vertical="center")
        pad = [None] * (span - 1)
        label = WriteOnlyCell(grid, value=note)      # spills right over empties
        label.font = Font(size=11, italic=True, color="FF666666")
        label.alignment = Alignment(vertical="center")
        return [c] + pad + [label]

    grid.append(control(1, "  ◀  frame number - type any value from 1 to "
                           f"{len(order)} here and the picture below repaints"))
    grid.append(control(f"=INDEX(Map!$A$1:${map_last}$1,n)",
                        "  ◀  which Data column draws it "
                        f"({len(order)} frames, {len(frames)} distinct - "
                        "held frames share a column)"))

    for r in range(h):
        row = []
        for c in range(w):
            i = r * w + c + 1                       # Data row for this cell
            cell = WriteOnlyCell(
                grid, value=f"=INDEX(Data!$A${i}:${data_last}${i},1,k)")
            cell.number_format = ";;;"              # keep the digits off the picture
            row.append(cell)
        grid.append(row)

    rng = f"A{GRID_TOP}:{last_col}{GRID_TOP + h - 1}"
    for i, (b, g, r) in enumerate(palette_bgr):
        rgb = f"FF{r:02X}{g:02X}{b:02X}"
        grid.conditional_formatting.add(rng, Rule(
            type="cellIs", operator="equal", formula=[str(i)], stopIfTrue=True,
            dxf=DifferentialStyle(
                fill=PatternFill(start_color=rgb, end_color=rgb, fill_type="solid"),
                font=Font(color=rgb))))          # belt and braces with ";;;"

    # --- Data / Map -------------------------------------------------------
    flat = np.stack([f.reshape(-1) for f in frames], axis=1)   # (w*h, ndistinct)
    for row in flat.tolist():
        data.append(row)
    fmap.append([i + 1 for i in order])                        # 1-based columns

    wb.defined_names.add(DefinedName("n", attr_text="Grid!$A$1"))
    wb.defined_names.add(DefinedName("k", attr_text="Grid!$A$2"))
    if title:
        wb.properties.title = title
    wb.save(path)

    px_w, px_h = geometry_px(cw, rh)
    return {"grid": (w, h), "cell_px": (px_w, px_h), "col_width": cw,
            "width_scale": width_scale,
            "row_height": rh, "distinct": len(frames), "frames": len(order),
            "rules": len(palette_bgr), "cells": w * h,
            "focus": f"{get_column_letter(max(1, w // 2))}{GRID_TOP + h // 2}"}
