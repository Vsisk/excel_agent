from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import random
from typing import Any, Literal, TypedDict


def random_8_digits() -> str:
    return f"{random.randint(0, 99_999_999):08d}"


def gen_id() -> str:
    return datetime.now().strftime("%Y%m%d") + random_8_digits()


class UsedRange(TypedDict):
    start_row: int
    end_row: int
    start_col: int
    end_col: int


class SheetInfoDict(TypedDict):
    sheet_id: str
    sheet_name: str
    sheet_index: int
    max_row: int
    max_col: int
    merged_cells: list[str]


class ExcelWorkbookDict(TypedDict):
    sheet_list: list[SheetInfoDict]


class SheetProfileDict(TypedDict):
    sheet_id: str
    used_range: UsedRange
    row_density: list[int]
    col_density: list[int]


class ClassificationDict(TypedDict):
    logic_area_type: Literal["fields", "fee_table", "detail_table", "plain_text", "unknown"]
    confidence: float


class ExcelParsePayload(TypedDict):
    excel_instance_id: str
    file_uri: str
    parse_mode: str


class ExcelParseRequest(TypedDict):
    request_type: str
    task_id: str
    site_id: str
    project_id: str
    payload: ExcelParsePayload


LogicAreaType = Literal["fields", "fee_table", "detail_table", "plain_text", "unknown"]


@dataclass(frozen=True)
class CellRange:
    start_row: int
    end_row: int
    start_col: int
    end_col: int

    def to_dict(self) -> UsedRange:
        return {
            "start_row": self.start_row,
            "end_row": self.end_row,
            "start_col": self.start_col,
            "end_col": self.end_col,
        }


@dataclass(frozen=True)
class SheetInfo:
    sheet_id: str
    sheet_name: str
    sheet_index: int
    max_row: int
    max_col: int
    merged_cells: list[str]

    def to_dict(self) -> SheetInfoDict:
        return {
            "sheet_id": self.sheet_id,
            "sheet_name": self.sheet_name,
            "sheet_index": self.sheet_index,
            "max_row": self.max_row,
            "max_col": self.max_col,
            "merged_cells": self.merged_cells,
        }


@dataclass(frozen=True)
class SheetProfile:
    sheet_id: str
    used_range: CellRange
    row_density: list[int]
    col_density: list[int]

    def to_dict(self) -> SheetProfileDict:
        return {
            "sheet_id": self.sheet_id,
            "used_range": self.used_range.to_dict(),
            "row_density": self.row_density,
            "col_density": self.col_density,
        }


@dataclass(frozen=True)
class ExcelRegion:
    sheet_id: str
    cell_range: CellRange
    raw_text: list[str]

    @property
    def row_count(self) -> int:
        return self.cell_range.end_row - self.cell_range.start_row + 1

    @property
    def col_count(self) -> int:
        return self.cell_range.end_col - self.cell_range.start_col + 1


JsonDict = dict[str, Any]
