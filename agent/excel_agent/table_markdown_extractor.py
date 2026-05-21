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
