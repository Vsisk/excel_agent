from __future__ import annotations

from typing import Any

from .excel_reader import ExcelReader
from .models import CellRange, SheetInfoDict, SheetProfile


class SheetProfiler:
    """Compute structural density signals for a sheet without materializing it."""

    @staticmethod
    def profile(reader: ExcelReader, sheet_info: SheetInfoDict) -> SheetProfile:
        max_row = sheet_info["max_row"]
        max_col = sheet_info["max_col"]
        row_density = [0 for _ in range(max_row)]
        col_density = [0 for _ in range(max_col)]

        start_row: int | None = None
        end_row = 0
        start_col: int | None = None
        end_col = 0

        for row_index, row in enumerate(
            reader.iter_rows(
                sheet_info["sheet_name"],
                min_row=1,
                max_row=max_row,
                min_col=1,
                max_col=max_col,
            ),
            start=1,
        ):
            non_empty_cols: list[int] = []
            for col_index, value in enumerate(row, start=1):
                if SheetProfiler._is_non_empty(value):
                    non_empty_cols.append(col_index)
                    col_density[col_index - 1] += 1

            row_density[row_index - 1] = len(non_empty_cols)
            if non_empty_cols:
                start_row = row_index if start_row is None else start_row
                end_row = row_index
                row_min_col = min(non_empty_cols)
                row_max_col = max(non_empty_cols)
                start_col = row_min_col if start_col is None else min(start_col, row_min_col)
                end_col = max(end_col, row_max_col)

        if start_row is None or start_col is None:
            used_range = CellRange(start_row=1, end_row=0, start_col=1, end_col=0)
        else:
            used_range = CellRange(
                start_row=start_row,
                end_row=end_row,
                start_col=start_col,
                end_col=end_col,
            )

        return SheetProfile(
            sheet_id=sheet_info["sheet_id"],
            used_range=used_range,
            row_density=row_density,
            col_density=col_density,
        )

    @staticmethod
    def _is_non_empty(value: Any) -> bool:
        if value is None:
            return False
        if isinstance(value, str):
            return value.strip() != ""
        return True
