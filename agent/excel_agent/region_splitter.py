from __future__ import annotations

from typing import Any

from .excel_reader import ExcelReader
from .models import CellRange, ExcelRegion, SheetInfoDict, SheetProfile


class RegionSplitter:
    """Split sheets into rectangular regions from density rules."""

    EMPTY_RUN_BREAK = 2
    DENSITY_JUMP_FACTOR = 3.0

    @classmethod
    def split(
        cls,
        reader: ExcelReader,
        sheet_info: SheetInfoDict,
        profile: SheetProfile,
    ) -> list[ExcelRegion]:
        used = profile.used_range
        if used.end_row < used.start_row or used.end_col < used.start_col:
            return []

        row_segments = cls._split_axis_by_empty_runs(
            profile.row_density,
            used.start_row,
            used.end_row,
        )
        row_segments = cls._split_row_segments_by_density(row_segments, profile.row_density)
        col_segments = cls._split_axis_by_empty_runs(
            profile.col_density,
            used.start_col,
            used.end_col,
        )

        regions: list[ExcelRegion] = []
        for row_start, row_end in row_segments:
            for col_start, col_end in col_segments:
                region = cls._region_from_candidate(
                    reader,
                    sheet_info["sheet_id"],
                    sheet_info["sheet_name"],
                    row_start,
                    row_end,
                    col_start,
                    col_end,
                )
                if region is not None:
                    regions.append(region)

        return regions

    @classmethod
    def _split_axis_by_empty_runs(
        cls,
        density: list[int],
        start: int,
        end: int,
    ) -> list[tuple[int, int]]:
        segments: list[tuple[int, int]] = []
        segment_start: int | None = None
        last_non_empty: int | None = None
        empty_run = 0

        for axis_index in range(start, end + 1):
            value = density[axis_index - 1] if axis_index - 1 < len(density) else 0
            if value == 0:
                empty_run += 1
                if segment_start is not None and empty_run >= cls.EMPTY_RUN_BREAK:
                    segments.append((segment_start, last_non_empty or axis_index - empty_run))
                    segment_start = None
                    last_non_empty = None
                continue

            if segment_start is None:
                segment_start = axis_index
            last_non_empty = axis_index
            empty_run = 0

        if segment_start is not None and last_non_empty is not None:
            segments.append((segment_start, last_non_empty))

        return segments

    @classmethod
    def _split_row_segments_by_density(
        cls,
        segments: list[tuple[int, int]],
        row_density: list[int],
    ) -> list[tuple[int, int]]:
        result: list[tuple[int, int]] = []
        for start, end in segments:
            current_start = start
            previous = row_density[start - 1]
            for row_index in range(start + 1, end + 1):
                current = row_density[row_index - 1]
                if cls._is_density_jump(previous, current):
                    result.append((current_start, row_index - 1))
                    current_start = row_index
                previous = current
            result.append((current_start, end))
        return result

    @classmethod
    def _is_density_jump(cls, previous: int, current: int) -> bool:
        if previous == 0 or current == 0:
            return False
        larger = max(previous, current)
        smaller = min(previous, current)
        return larger >= 4 and larger / smaller >= cls.DENSITY_JUMP_FACTOR

    @classmethod
    def _region_from_candidate(
        cls,
        reader: ExcelReader,
        sheet_id: str,
        sheet_name: str,
        start_row: int,
        end_row: int,
        start_col: int,
        end_col: int,
    ) -> ExcelRegion | None:
        trimmed_start_row: int | None = None
        trimmed_end_row = 0
        trimmed_start_col: int | None = None
        trimmed_end_col = 0

        for row_offset, row in enumerate(
            reader.iter_rows(
                sheet_name,
                min_row=start_row,
                max_row=end_row,
                min_col=start_col,
                max_col=end_col,
            )
        ):
            row_index = start_row + row_offset
            non_empty_cols = [
                start_col + col_offset
                for col_offset, value in enumerate(row)
                if cls._is_non_empty(value)
            ]
            if not non_empty_cols:
                continue
            trimmed_start_row = row_index if trimmed_start_row is None else trimmed_start_row
            trimmed_end_row = row_index
            row_min_col = min(non_empty_cols)
            row_max_col = max(non_empty_cols)
            trimmed_start_col = row_min_col if trimmed_start_col is None else min(trimmed_start_col, row_min_col)
            trimmed_end_col = max(trimmed_end_col, row_max_col)

        if trimmed_start_row is None or trimmed_start_col is None:
            return None

        raw_text = cls._raw_text(
            reader,
            sheet_name,
            trimmed_start_row,
            trimmed_end_row,
            trimmed_start_col,
            trimmed_end_col,
        )
        return ExcelRegion(
            sheet_id=sheet_id,
            cell_range=CellRange(
                trimmed_start_row,
                trimmed_end_row,
                trimmed_start_col,
                trimmed_end_col,
            ),
            raw_text=raw_text,
        )

    @classmethod
    def _raw_text(
        cls,
        reader: ExcelReader,
        sheet_name: str,
        start_row: int,
        end_row: int,
        start_col: int,
        end_col: int,
    ) -> list[str]:
        lines: list[str] = []
        for row in reader.iter_rows(
            sheet_name,
            min_row=start_row,
            max_row=end_row,
            min_col=start_col,
            max_col=end_col,
        ):
            values = [cls._stringify(value) for value in row if cls._is_non_empty(value)]
            if values:
                lines.append(" ".join(values))
        return lines

    @staticmethod
    def _is_non_empty(value: Any) -> bool:
        return value is not None and (not isinstance(value, str) or value.strip() != "")

    @staticmethod
    def _stringify(value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()
