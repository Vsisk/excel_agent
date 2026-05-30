from __future__ import annotations

from typing import Any

from openpyxl.utils.cell import coordinate_to_tuple

from .excel_reader import ExcelReader
from .models import CellRange, ExcelRegion
from .region_markdown_builder import RegionMarkdownBuilder
from .region_splitter import RegionSplitter
from .sheet_profiler import SheetProfiler


CellCoordinate = str | tuple[int, int]
BBox = list[int] | tuple[int, int, int, int]
OutputCellRange = dict[str, int]


def get_table_md_by_cell(
    file_path: str,
    *,
    sheet_number: int,
    cell_coordinate: CellCoordinate,
) -> str:
    """Return table markdown for the rule-split region containing a cell.

    sheet_number is 1-based. cell_coordinate accepts Excel A1 notation, such as
    "B6", or a (row, column) tuple using 1-based indexes.
    """
    row, col = _normalize_cell_coordinate(cell_coordinate)
    return get_table_md_by_bbox(
        file_path,
        sheet_number=sheet_number,
        bbox=[col, row, col, row],
    )


def get_table_md_by_bbox(
    file_path: str,
    *,
    sheet_number: int,
    bbox: BBox,
) -> str:
    """Return table markdown for rule-split regions intersecting a bbox.

    bbox order is [left, top, right, bottom], using 1-based Excel column and row
    indexes.
    """
    target_range = _normalize_bbox(bbox)
    reader = ExcelReader(file_path)
    try:
        workbook = reader.read()
        sheet_info = _sheet_by_number(workbook["sheet_list"], sheet_number)
        profile = SheetProfiler.profile(reader, sheet_info)
        regions = RegionSplitter.split(reader, sheet_info, profile)
        matched_regions = _find_regions_intersecting_range(regions, target_range)
        if not matched_regions:
            return ""

        markdown_parts = []
        for index, region in enumerate(matched_regions, start=1):
            rows = reader.read_range(sheet_info["sheet_name"], region.cell_range)
            snapshot = RegionMarkdownBuilder.build_region_snapshot(
                region_id=f"region_{index}",
                region=region,
                rows=rows,
            )
            markdown_parts.append(snapshot.markdown)
        return "\n".join(part for part in markdown_parts if part)
    finally:
        reader.close()


def get_table_md_by_cell_range(
    file_path: str,
    *,
    sheet_number: int,
    cell_range: OutputCellRange,
) -> str:
    """Return markdown for an exact 0-based half-open output cell range.

    Accepted shapes:
    - {"start_row": 0, "end_row": 2, "start_col": 0, "end_col": 2}
    - {"left": 0, "right": 2, "top": 0, "bottom": 2}
    """
    target_range = _normalize_output_cell_range(cell_range)
    reader = ExcelReader(file_path)
    try:
        workbook = reader.read()
        sheet_info = _sheet_by_number(workbook["sheet_list"], sheet_number)
        rows = reader.read_range(sheet_info["sheet_name"], target_range)
        snapshot = RegionMarkdownBuilder.build_region_snapshot(
            region_id="range_1",
            region=ExcelRegion(sheet_info["sheet_id"], target_range, raw_text=[]),
            rows=rows,
        )
        return snapshot.markdown
    finally:
        reader.close()


def _normalize_cell_coordinate(cell_coordinate: CellCoordinate) -> tuple[int, int]:
    if isinstance(cell_coordinate, str):
        return coordinate_to_tuple(cell_coordinate)

    if (
        isinstance(cell_coordinate, tuple)
        and len(cell_coordinate) == 2
        and all(isinstance(value, int) for value in cell_coordinate)
    ):
        row, col = cell_coordinate
        if row < 1 or col < 1:
            raise ValueError("cell_coordinate row and column must be 1-based positive integers")
        return row, col

    raise ValueError("cell_coordinate must be A1 notation or a (row, column) tuple")


def _normalize_output_cell_range(cell_range: OutputCellRange) -> CellRange:
    if not isinstance(cell_range, dict):
        raise ValueError("cell_range must be a dict")

    if {"start_row", "end_row", "start_col", "end_col"} <= set(cell_range):
        start_row = cell_range["start_row"]
        end_row = cell_range["end_row"]
        start_col = cell_range["start_col"]
        end_col = cell_range["end_col"]
    elif {"left", "right", "top", "bottom"} <= set(cell_range):
        start_row = cell_range["top"]
        end_row = cell_range["bottom"]
        start_col = cell_range["left"]
        end_col = cell_range["right"]
    else:
        raise ValueError(
            "cell_range must contain start_row/end_row/start_col/end_col or left/right/top/bottom"
        )

    if not all(isinstance(value, int) for value in [start_row, end_row, start_col, end_col]):
        raise ValueError("cell_range values must be integers")
    if start_row < 0 or start_col < 0 or end_row <= start_row or end_col <= start_col:
        raise ValueError("cell_range must be 0-based half-open coordinates with positive width and height")

    return CellRange(
        start_row=start_row + 1,
        end_row=end_row,
        start_col=start_col + 1,
        end_col=end_col,
    )


def _normalize_bbox(bbox: BBox) -> CellRange:
    if not isinstance(bbox, (list, tuple)) or len(bbox) != 4:
        raise ValueError("bbox must be [left, top, right, bottom]")
    left, top, right, bottom = bbox
    if not all(isinstance(value, int) for value in [left, top, right, bottom]):
        raise ValueError("bbox values must be integers")
    if left < 1 or top < 1 or right < left or bottom < top:
        raise ValueError("bbox must use 1-based [left, top, right, bottom] with right>=left and bottom>=top")
    return CellRange(start_row=top, end_row=bottom, start_col=left, end_col=right)


def _sheet_by_number(sheet_list: list[dict[str, Any]], sheet_number: int) -> dict[str, Any]:
    if sheet_number < 1 or sheet_number > len(sheet_list):
        raise ValueError(f"sheet_number must be between 1 and {len(sheet_list)}")
    return sheet_list[sheet_number - 1]


def _find_regions_intersecting_range(
    regions: list[ExcelRegion],
    target_range: CellRange,
) -> list[ExcelRegion]:
    return [
        region
        for region in regions
        if _ranges_intersect(region.cell_range, target_range)
    ]


def _ranges_intersect(left: CellRange, right: CellRange) -> bool:
    return (
        left.start_row <= right.end_row
        and right.start_row <= left.end_row
        and left.start_col <= right.end_col
        and right.start_col <= left.end_col
    )
