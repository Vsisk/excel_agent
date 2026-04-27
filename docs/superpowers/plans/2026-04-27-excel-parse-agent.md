# Excel Parse Agent Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the Agent-side `EXCEL_PARSE` pipeline that reads Excel workbooks with `openpyxl`, splits sheets into logical regions, classifies regions, and returns `logic_page_list`, `logic_area_list`, and `parse_index` through the WebSocket-style handler contract.

**Architecture:** The implementation is a small pure-Python package under `agent/excel_agent/`. `excel_reader.py` owns workbook metadata and lightweight worksheet access, `sheet_profiler.py` computes densities from streamed rows, `region_splitter.py` performs deterministic area segmentation, `region_classifier.py` applies rule-based classification with an LLM-ready seam, `logic_builder.py` emits the final logic JSON, and `excel_parse_handler.py` orchestrates the request-to-response flow.

**Tech Stack:** Python 3.10+, `openpyxl` for Excel reading, `dataclasses` and `TypedDict` for clear structures, `pytest` for tests. Do not use pandas.

---

## File Structure

- Create: `agent/__init__.py`
  - Makes `agent` importable.
- Create: `agent/excel_agent/__init__.py`
  - Exposes the Excel parse package.
- Create: `agent/excel_agent/models.py`
  - Defines `TypedDict` structures, frozen dataclasses for in-memory objects, and `gen_id()`.
- Create: `agent/excel_agent/excel_reader.py`
  - Opens workbooks in `read_only=True, data_only=True`, extracts sheet metadata, and exposes controlled row iterators without loading whole sheets.
- Create: `agent/excel_agent/sheet_profiler.py`
  - Computes `used_range`, `row_density`, and `col_density` from row iteration.
- Create: `agent/excel_agent/region_splitter.py`
  - Splits regions using consecutive empty rows, consecutive empty columns, and row density jumps.
- Create: `agent/excel_agent/region_classifier.py`
  - Classifies regions as `fields`, `fee_table`, `detail_table`, `plain_text`, or `unknown` using rules only.
- Create: `agent/excel_agent/logic_builder.py`
  - Builds `logic_page` and `logic_area` dictionaries.
- Create: `agent/excel_agent/excel_parse_handler.py`
  - Validates and handles `EXCEL_PARSE` requests, returns `EXCEL_PARSE_RESULT`.
- Create: `tests/test_excel_parse_agent.py`
  - Covers multi-sheet parsing, region splitting, classification, output shape, and ID/range expectations.

---

### Task 1: Package Skeleton And Shared Models

**Files:**
- Create: `agent/__init__.py`
- Create: `agent/excel_agent/__init__.py`
- Create: `agent/excel_agent/models.py`
- Test: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Write the failing tests for IDs and model imports**

Create `tests/test_excel_parse_agent.py` with:

```python
import re

from agent.excel_agent.models import CellRange, ExcelRegion, gen_id


def test_gen_id_uses_date_and_eight_digits():
    value = gen_id()

    assert re.fullmatch(r"\d{16}", value)


def test_region_model_uses_one_based_range():
    region = ExcelRegion(
        sheet_id="sheet_1",
        cell_range=CellRange(start_row=1, end_row=3, start_col=1, end_col=2),
        raw_text=["A B", "C D"],
    )

    assert region.cell_range.start_row == 1
    assert region.cell_range.start_col == 1
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py
```

Expected: FAIL with `ModuleNotFoundError: No module named 'agent'`.

- [ ] **Step 3: Create package skeleton**

Create `agent/__init__.py`:

```python
"""Agent package."""
```

Create `agent/excel_agent/__init__.py`:

```python
"""Excel parsing agent package."""
```

- [ ] **Step 4: Implement shared models**

Create `agent/excel_agent/models.py`:

```python
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
```

- [ ] **Step 5: Run test to verify it passes**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py
```

Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add agent tests/test_excel_parse_agent.py
git commit -m "feat: add excel agent models"
```

If this workspace is not a git repository, skip the commit and record that in the implementation notes.

---

### Task 2: Excel Reader

**Files:**
- Create: `agent/excel_agent/excel_reader.py`
- Modify: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Add failing reader test**

Append to `tests/test_excel_parse_agent.py`:

```python
from pathlib import Path

from openpyxl import Workbook

from agent.excel_agent.excel_reader import ExcelReader


def make_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "Invoice"
    ws["B1"] = "Amount"
    ws["A2"] = "INV-001"
    ws["B2"] = 120
    detail = wb.create_sheet("Details")
    detail["A1"] = "Item"
    detail["B1"] = "Qty"
    detail["A2"] = "Storage"
    detail["B2"] = 3
    wb.save(path)


def test_excel_reader_reads_sheet_metadata_and_rows(tmp_path):
    path = tmp_path / "sample.xlsx"
    make_workbook(path)

    reader = ExcelReader(str(path))
    workbook = reader.read()

    assert [sheet["sheet_name"] for sheet in workbook["sheet_list"]] == ["Summary", "Details"]
    assert workbook["sheet_list"][0]["sheet_index"] == 0
    assert workbook["sheet_list"][0]["max_row"] == 2
    assert workbook["sheet_list"][0]["max_col"] == 2
    rows = list(reader.iter_rows("Summary", min_row=1, max_row=2, min_col=1, max_col=2))
    reader.close()

    assert rows == [["Invoice", "Amount"], ["INV-001", 120]]
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py::test_excel_reader_reads_sheet_metadata_and_rows
```

Expected: FAIL with `ModuleNotFoundError` for `excel_reader`.

- [ ] **Step 3: Implement reader**

Create `agent/excel_agent/excel_reader.py`:

```python
from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from .models import ExcelWorkbookDict, SheetInfo


class ExcelReader:
    def __init__(self, file_path: str):
        self.file_path = str(Path(file_path))
        self._workbook = None
        self._sheet_info_by_name: dict[str, SheetInfo] = {}

    def _ensure_open(self):
        if self._workbook is None:
            self._workbook = load_workbook(self.file_path, read_only=True, data_only=True)
        return self._workbook

    def read(self) -> ExcelWorkbookDict:
        workbook = self._ensure_open()
        sheet_list: list[dict] = []
        self._sheet_info_by_name.clear()

        for index, sheet_name in enumerate(workbook.sheetnames):
            sheet = workbook[sheet_name]
            info = SheetInfo(
                sheet_id=f"sheet_{index + 1}",
                sheet_name=sheet_name,
                sheet_index=index,
                max_row=sheet.max_row or 0,
                max_col=sheet.max_column or 0,
                merged_cells=self._merged_ranges(sheet),
            )
            self._sheet_info_by_name[sheet_name] = info
            sheet_list.append(info.to_dict())

        return {"sheet_list": sheet_list}

    def get_sheet_info(self, sheet_name: str) -> SheetInfo:
        if not self._sheet_info_by_name:
            self.read()
        return self._sheet_info_by_name[sheet_name]

    def iter_rows(
        self,
        sheet_name: str,
        *,
        min_row: int = 1,
        max_row: int | None = None,
        min_col: int = 1,
        max_col: int | None = None,
    ) -> Iterator[list[Any]]:
        workbook = self._ensure_open()
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows(
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            values_only=True,
        ):
            yield list(row)

    def close(self) -> None:
        if self._workbook is not None:
            self._workbook.close()
            self._workbook = None

    @staticmethod
    def _merged_ranges(sheet: Worksheet) -> list[str]:
        merged = getattr(sheet, "merged_cells", None)
        ranges = getattr(merged, "ranges", []) if merged is not None else []
        return [str(cell_range) for cell_range in ranges]
```

- [ ] **Step 4: Run reader test**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py::test_excel_reader_reads_sheet_metadata_and_rows
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add agent/excel_agent/excel_reader.py tests/test_excel_parse_agent.py
git commit -m "feat: add streaming excel reader"
```

If this workspace is not a git repository, skip the commit and record that in the implementation notes.

---

### Task 3: Sheet Profiler

**Files:**
- Create: `agent/excel_agent/sheet_profiler.py`
- Modify: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Add failing profiler test**

Append to `tests/test_excel_parse_agent.py`:

```python
from agent.excel_agent.sheet_profiler import SheetProfiler


def test_sheet_profiler_computes_used_range_and_density(tmp_path):
    path = tmp_path / "profile.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Profile"
    ws["B2"] = "Name"
    ws["C2"] = "Amount"
    ws["B3"] = "A"
    ws["C3"] = 10
    ws["B6"] = "Footer"
    wb.save(path)

    reader = ExcelReader(str(path))
    workbook = reader.read()
    sheet_info = workbook["sheet_list"][0]

    profile = SheetProfiler.profile(reader, sheet_info)
    reader.close()

    assert profile.used_range.to_dict() == {
        "start_row": 2,
        "end_row": 6,
        "start_col": 2,
        "end_col": 3,
    }
    assert profile.row_density == [0, 2, 2, 0, 0, 1]
    assert profile.col_density == [0, 3, 2]
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py::test_sheet_profiler_computes_used_range_and_density
```

Expected: FAIL with `ModuleNotFoundError` for `sheet_profiler`.

- [ ] **Step 3: Implement profiler**

Create `agent/excel_agent/sheet_profiler.py`:

```python
from __future__ import annotations

from typing import Any

from .excel_reader import ExcelReader
from .models import CellRange, SheetInfoDict, SheetProfile


class SheetProfiler:
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
            non_empty_cols = []
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
```

- [ ] **Step 4: Run profiler test**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py::test_sheet_profiler_computes_used_range_and_density
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add agent/excel_agent/sheet_profiler.py tests/test_excel_parse_agent.py
git commit -m "feat: profile excel sheets"
```

If this workspace is not a git repository, skip the commit and record that in the implementation notes.

---

### Task 4: Region Splitter

**Files:**
- Create: `agent/excel_agent/region_splitter.py`
- Modify: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Add failing splitter test**

Append to `tests/test_excel_parse_agent.py`:

```python
from agent.excel_agent.region_splitter import RegionSplitter


def test_region_splitter_splits_by_empty_rows_and_columns(tmp_path):
    path = tmp_path / "regions.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Regions"
    ws["A1"] = "Customer"
    ws["B1"] = "ACME"
    ws["A2"] = "Billing Month"
    ws["B2"] = "2026-04"
    ws["A5"] = "Item"
    ws["B5"] = "Qty"
    ws["C5"] = "Amount"
    ws["A6"] = "Compute"
    ws["B6"] = 2
    ws["C6"] = 100
    ws["F5"] = "Notes"
    ws["F6"] = "Long running text"
    wb.save(path)

    reader = ExcelReader(str(path))
    sheet_info = reader.read()["sheet_list"][0]
    profile = SheetProfiler.profile(reader, sheet_info)

    regions = RegionSplitter.split(reader, sheet_info, profile)
    reader.close()

    assert [region.cell_range.to_dict() for region in regions] == [
        {"start_row": 1, "end_row": 2, "start_col": 1, "end_col": 2},
        {"start_row": 5, "end_row": 6, "start_col": 1, "end_col": 3},
        {"start_row": 5, "end_row": 6, "start_col": 6, "end_col": 6},
    ]
    assert regions[1].raw_text == ["Item Qty Amount", "Compute 2 100"]
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py::test_region_splitter_splits_by_empty_rows_and_columns
```

Expected: FAIL with `ModuleNotFoundError` for `region_splitter`.

- [ ] **Step 3: Implement splitter**

Create `agent/excel_agent/region_splitter.py`:

```python
from __future__ import annotations

from typing import Any

from .excel_reader import ExcelReader
from .models import CellRange, ExcelRegion, SheetInfoDict, SheetProfile


class RegionSplitter:
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
                raw_text = cls._raw_text(
                    reader,
                    sheet_info["sheet_name"],
                    row_start,
                    row_end,
                    col_start,
                    col_end,
                )
                if raw_text:
                    regions.append(
                        ExcelRegion(
                            sheet_id=sheet_info["sheet_id"],
                            cell_range=CellRange(row_start, row_end, col_start, col_end),
                            raw_text=raw_text,
                        )
                    )

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
```

- [ ] **Step 4: Run splitter test**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py::test_region_splitter_splits_by_empty_rows_and_columns
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add agent/excel_agent/region_splitter.py tests/test_excel_parse_agent.py
git commit -m "feat: split excel sheets into regions"
```

If this workspace is not a git repository, skip the commit and record that in the implementation notes.

---

### Task 5: Region Classifier

**Files:**
- Create: `agent/excel_agent/region_classifier.py`
- Modify: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Add failing classifier tests**

Append to `tests/test_excel_parse_agent.py`:

```python
from agent.excel_agent.region_classifier import RegionClassifier


def test_region_classifier_detects_fields():
    region = ExcelRegion(
        sheet_id="sheet_1",
        cell_range=CellRange(1, 2, 1, 2),
        raw_text=["Customer ACME", "Billing Month 2026-04"],
    )

    assert RegionClassifier.classify(region) == {
        "logic_area_type": "fields",
        "confidence": 0.82,
    }


def test_region_classifier_detects_tables_and_plain_text():
    table = ExcelRegion(
        sheet_id="sheet_1",
        cell_range=CellRange(1, 3, 1, 3),
        raw_text=["Item Qty Amount", "Compute 2 100", "Storage 3 50"],
    )
    text = ExcelRegion(
        sheet_id="sheet_1",
        cell_range=CellRange(1, 1, 1, 1),
        raw_text=["This billing description contains a long explanatory paragraph with many words."],
    )

    assert RegionClassifier.classify(table)["logic_area_type"] == "fee_table"
    assert RegionClassifier.classify(text)["logic_area_type"] == "plain_text"
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py::test_region_classifier_detects_fields tests/test_excel_parse_agent.py::test_region_classifier_detects_tables_and_plain_text
```

Expected: FAIL with `ModuleNotFoundError` for `region_classifier`.

- [ ] **Step 3: Implement classifier**

Create `agent/excel_agent/region_classifier.py`:

```python
from __future__ import annotations

import re

from .models import ClassificationDict, ExcelRegion


class RegionClassifier:
    HEADER_KEYWORDS = {
        "amount",
        "fee",
        "price",
        "qty",
        "quantity",
        "item",
        "description",
        "subtotal",
        "total",
        "金额",
        "费用",
        "数量",
        "项目",
        "名称",
    }

    @classmethod
    def classify(cls, region: ExcelRegion) -> ClassificationDict:
        if region.row_count <= 5 and cls._looks_like_fields(region):
            return {"logic_area_type": "fields", "confidence": 0.82}

        if cls._has_clear_header(region):
            area_type = "fee_table" if cls._looks_like_fee_table(region) else "detail_table"
            return {"logic_area_type": area_type, "confidence": 0.78}

        if cls._looks_like_plain_text(region):
            return {"logic_area_type": "plain_text", "confidence": 0.74}

        return {"logic_area_type": "unknown", "confidence": 0.35}

    @classmethod
    def _looks_like_fields(cls, region: ExcelRegion) -> bool:
        if region.col_count < 2 or not region.raw_text:
            return False
        short_lines = [line for line in region.raw_text if 2 <= len(line.split()) <= 8]
        return len(short_lines) >= max(1, len(region.raw_text) // 2)

    @classmethod
    def _has_clear_header(cls, region: ExcelRegion) -> bool:
        if region.row_count < 2 or region.col_count < 2 or not region.raw_text:
            return False
        first_line = region.raw_text[0].lower()
        matches = sum(1 for keyword in cls.HEADER_KEYWORDS if keyword in first_line)
        return matches >= 2 or (matches >= 1 and region.col_count >= 3)

    @classmethod
    def _looks_like_fee_table(cls, region: ExcelRegion) -> bool:
        joined = " ".join(region.raw_text).lower()
        return any(token in joined for token in ["amount", "fee", "price", "total", "金额", "费用"])

    @classmethod
    def _looks_like_plain_text(cls, region: ExcelRegion) -> bool:
        joined = " ".join(region.raw_text)
        words = re.findall(r"\S+", joined)
        return region.col_count <= 2 and len(words) >= 10
```

- [ ] **Step 4: Run classifier tests**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py::test_region_classifier_detects_fields tests/test_excel_parse_agent.py::test_region_classifier_detects_tables_and_plain_text
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add agent/excel_agent/region_classifier.py tests/test_excel_parse_agent.py
git commit -m "feat: classify excel regions"
```

If this workspace is not a git repository, skip the commit and record that in the implementation notes.

---

### Task 6: Logic Builder

**Files:**
- Create: `agent/excel_agent/logic_builder.py`
- Modify: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Add failing builder test**

Append to `tests/test_excel_parse_agent.py`:

```python
from agent.excel_agent.logic_builder import LogicBuilder


def test_logic_builder_creates_page_and_area():
    sheet_info = {
        "sheet_id": "sheet_1",
        "sheet_name": "Summary",
        "sheet_index": 0,
        "max_row": 2,
        "max_col": 2,
        "merged_cells": [],
    }
    region = ExcelRegion(
        sheet_id="sheet_1",
        cell_range=CellRange(1, 2, 1, 2),
        raw_text=["Customer ACME", "Billing Month 2026-04"],
    )

    page = LogicBuilder.build_page("excel_1", sheet_info)
    area = LogicBuilder.build_area("excel_1", sheet_info, region, {"logic_area_type": "fields", "confidence": 0.82})

    assert page["logic_page_relation"] == {
        "type": "excel",
        "excel_instance_id": "excel_1",
        "sheet_id": "sheet_1",
        "sheet_name": "Summary",
        "sheet_index": 0,
    }
    assert area["logic_area_type"] == "fields"
    assert area["location_list"][0]["cell_range"] == {
        "start_row": 1,
        "end_row": 2,
        "start_col": 1,
        "end_col": 2,
    }
    assert area["location_list"][0]["raw_excel_text_list"] == region.raw_text
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py::test_logic_builder_creates_page_and_area
```

Expected: FAIL with `ModuleNotFoundError` for `logic_builder`.

- [ ] **Step 3: Implement builder**

Create `agent/excel_agent/logic_builder.py`:

```python
from __future__ import annotations

from .models import ClassificationDict, ExcelRegion, JsonDict, SheetInfoDict, gen_id


class LogicBuilder:
    @staticmethod
    def build_page(excel_instance_id: str, sheet_info: SheetInfoDict) -> JsonDict:
        return {
            "logic_page_id": gen_id(),
            "logic_page_relation": {
                "type": "excel",
                "excel_instance_id": excel_instance_id,
                "sheet_id": sheet_info["sheet_id"],
                "sheet_name": sheet_info["sheet_name"],
                "sheet_index": sheet_info["sheet_index"],
            },
        }

    @staticmethod
    def build_area(
        excel_instance_id: str,
        sheet_info: SheetInfoDict,
        region: ExcelRegion,
        classification: ClassificationDict,
    ) -> JsonDict:
        cell_range = region.cell_range.to_dict()
        area_type = classification["logic_area_type"]
        return {
            "logic_area_id": gen_id(),
            "logic_area_name": f"{sheet_info['sheet_name']}!R{cell_range['start_row']}C{cell_range['start_col']}",
            "logic_area_type": area_type,
            "logic_area_description": f"{area_type} region from sheet {sheet_info['sheet_name']}",
            "location_list": [
                {
                    "type": "excel",
                    "excel_instance_id": excel_instance_id,
                    "sheet_id": sheet_info["sheet_id"],
                    "sheet_index": sheet_info["sheet_index"],
                    "cell_range": cell_range,
                    "raw_excel_text_list": region.raw_text,
                }
            ],
            "classification": classification,
        }
```

- [ ] **Step 4: Run builder test**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py::test_logic_builder_creates_page_and_area
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add agent/excel_agent/logic_builder.py tests/test_excel_parse_agent.py
git commit -m "feat: build excel logic output"
```

If this workspace is not a git repository, skip the commit and record that in the implementation notes.

---

### Task 7: Excel Parse Handler

**Files:**
- Create: `agent/excel_agent/excel_parse_handler.py`
- Modify: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Add failing end-to-end handler test**

Append to `tests/test_excel_parse_agent.py`:

```python
from agent.excel_agent.excel_parse_handler import handle_excel_parse


def test_handle_excel_parse_returns_ws_result_for_multiple_sheets(tmp_path):
    path = tmp_path / "full.xlsx"
    make_workbook(path)

    response = handle_excel_parse(
        {
            "request_type": "EXCEL_PARSE",
            "task_id": "task_1",
            "site_id": "site_1",
            "project_id": "project_1",
            "payload": {
                "excel_instance_id": "excel_1",
                "file_uri": str(path),
                "parse_mode": "full",
            },
        }
    )

    assert response["request_type"] == "EXCEL_PARSE_RESULT"
    assert response["task_id"] == "task_1"
    assert response["status"] == "success"
    assert len(response["payload"]["logic_page_list"]) == 2
    assert len(response["payload"]["logic_area_list"]) >= 2
    assert response["payload"]["parse_index"]["excel_instance_id"] == "excel_1"
    assert response["payload"]["parse_index"]["sheet_count"] == 2
    assert response["payload"]["parse_index"]["area_count"] == len(response["payload"]["logic_area_list"])
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py::test_handle_excel_parse_returns_ws_result_for_multiple_sheets
```

Expected: FAIL with `ModuleNotFoundError` for `excel_parse_handler`.

- [ ] **Step 3: Implement handler**

Create `agent/excel_agent/excel_parse_handler.py`:

```python
from __future__ import annotations

from .excel_reader import ExcelReader
from .logic_builder import LogicBuilder
from .models import ExcelParseRequest, JsonDict
from .region_classifier import RegionClassifier
from .region_splitter import RegionSplitter
from .sheet_profiler import SheetProfiler


def handle_excel_parse(req: ExcelParseRequest) -> JsonDict:
    _validate_request(req)
    payload = req["payload"]
    excel_instance_id = payload["excel_instance_id"]
    reader = ExcelReader(payload["file_uri"])

    logic_page_list: list[JsonDict] = []
    logic_area_list: list[JsonDict] = []
    sheet_profiles: dict[str, JsonDict] = {}

    try:
        workbook = reader.read()
        for sheet_info in workbook["sheet_list"]:
            profile = SheetProfiler.profile(reader, sheet_info)
            sheet_profiles[sheet_info["sheet_id"]] = profile.to_dict()
            logic_page_list.append(LogicBuilder.build_page(excel_instance_id, sheet_info))

            regions = RegionSplitter.split(reader, sheet_info, profile)
            for region in regions:
                classification = RegionClassifier.classify(region)
                logic_area_list.append(
                    LogicBuilder.build_area(
                        excel_instance_id,
                        sheet_info,
                        region,
                        classification,
                    )
                )
    finally:
        reader.close()

    return {
        "request_type": "EXCEL_PARSE_RESULT",
        "task_id": req["task_id"],
        "status": "success",
        "payload": {
            "logic_page_list": logic_page_list,
            "logic_area_list": logic_area_list,
            "parse_index": {
                "excel_instance_id": excel_instance_id,
                "sheet_count": len(logic_page_list),
                "area_count": len(logic_area_list),
                "sheet_profiles": sheet_profiles,
            },
        },
    }


def _validate_request(req: ExcelParseRequest) -> None:
    if req.get("request_type") != "EXCEL_PARSE":
        raise ValueError("request_type must be EXCEL_PARSE")
    payload = req.get("payload")
    if not isinstance(payload, dict):
        raise ValueError("payload is required")
    for key in ["excel_instance_id", "file_uri", "parse_mode"]:
        if not payload.get(key):
            raise ValueError(f"payload.{key} is required")
    if payload["parse_mode"] != "full":
        raise ValueError("only parse_mode=full is supported")
```

- [ ] **Step 4: Run handler test**

Run:

```bash
python -m pytest -q tests/test_excel_parse_agent.py::test_handle_excel_parse_returns_ws_result_for_multiple_sheets
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add agent/excel_agent/excel_parse_handler.py tests/test_excel_parse_agent.py
git commit -m "feat: handle excel parse requests"
```

If this workspace is not a git repository, skip the commit and record that in the implementation notes.

---

### Task 8: Full Verification And Edge Cases

**Files:**
- Modify: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Add edge-case tests for validation and empty sheets**

Append to `tests/test_excel_parse_agent.py`:

```python
import pytest


def test_handle_excel_parse_rejects_wrong_request_type(tmp_path):
    path = tmp_path / "bad.xlsx"
    make_workbook(path)

    with pytest.raises(ValueError, match="request_type must be EXCEL_PARSE"):
        handle_excel_parse(
            {
                "request_type": "OTHER",
                "task_id": "task_1",
                "site_id": "site_1",
                "project_id": "project_1",
                "payload": {
                    "excel_instance_id": "excel_1",
                    "file_uri": str(path),
                    "parse_mode": "full",
                },
            }
        )


def test_empty_sheet_returns_page_without_area(tmp_path):
    path = tmp_path / "empty.xlsx"
    wb = Workbook()
    wb.active.title = "Empty"
    wb.save(path)

    response = handle_excel_parse(
        {
            "request_type": "EXCEL_PARSE",
            "task_id": "task_empty",
            "site_id": "site_1",
            "project_id": "project_1",
            "payload": {
                "excel_instance_id": "excel_empty",
                "file_uri": str(path),
                "parse_mode": "full",
            },
        }
    )

    assert response["status"] == "success"
    assert len(response["payload"]["logic_page_list"]) == 1
    assert response["payload"]["logic_area_list"] == []
```

- [ ] **Step 2: Run full test suite**

Run:

```bash
python -m pytest -q
```

Expected: all tests PASS.

- [ ] **Step 3: Run compile check**

Run:

```bash
python -m py_compile agent/excel_agent/models.py agent/excel_agent/excel_reader.py agent/excel_agent/sheet_profiler.py agent/excel_agent/region_splitter.py agent/excel_agent/region_classifier.py agent/excel_agent/logic_builder.py agent/excel_agent/excel_parse_handler.py
```

Expected: no output and exit code 0.

- [ ] **Step 4: Review constraints manually**

Confirm:

```text
openpyxl is used for Excel parsing.
pandas is not imported anywhere.
load_workbook uses read_only=True and data_only=True.
All emitted ranges use 1-based row and column indexes.
The handler input request_type is EXCEL_PARSE.
The handler output request_type is EXCEL_PARSE_RESULT.
No HTTP or frontend code was added.
No real LLM call was added.
```

- [ ] **Step 5: Commit**

```bash
git add agent tests/test_excel_parse_agent.py
git commit -m "test: verify excel parse agent flow"
```

If this workspace is not a git repository, skip the commit and record that in the implementation notes.

---

## Self-Review

**Spec coverage:** This plan covers the required `agent/excel_agent/` module structure, `openpyxl` read-only workbook loading, multi-sheet metadata extraction, sheet profiling, rule-based region splitting, rule-based classification with no LLM call, logic page and logic area output, and a complete `EXCEL_PARSE` request to `EXCEL_PARSE_RESULT` response path.

**Placeholder scan:** The plan contains no `TBD`, `TODO`, `implement later`, or vague test instructions. Each task includes concrete files, test code, implementation code, commands, and expected outcomes.

**Type consistency:** The plan defines `CellRange`, `SheetInfoDict`, `SheetProfile`, `ExcelRegion`, `ClassificationDict`, `ExcelParseRequest`, and `JsonDict` before later tasks reference them. Later modules use the same class and function names: `ExcelReader`, `SheetProfiler.profile`, `RegionSplitter.split`, `RegionClassifier.classify`, `LogicBuilder.build_page`, `LogicBuilder.build_area`, and `handle_excel_parse`.
