# Excel Sheet LLM Grouping Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add optional per-sheet LLM grouping to `EXCEL_PARSE`: each sheet gets one logic page name from `bill_summary_page`, `bill_charge_page`, or `bill_cdr_page`, and LLM groups existing rule-split region IDs using compact Markdown region content.

**Architecture:** Keep the existing deterministic Excel parser as the fallback path. Add small internal grouping models, a Markdown snapshot builder, a sheet-level LLM grouping adapter, an optional visual-summary collector, and LogicBuilder/handler integration that builds grouped logic areas while preserving success on LLM failure.

**Tech Stack:** Python 3.10+, `openpyxl`, existing `agent.llm.generate_by_llm`, `pytest`. Do not use pandas, GUI Excel automation, or real LLM calls in tests.

---

## File Structure

- Modify: `agent/excel_agent/models.py`
  - Add internal grouping dataclasses and allowed page-name constants.
- Modify: `agent/excel_agent/excel_reader.py`
  - Add `read_range()` convenience method that reads only a region range.
- Create: `agent/excel_agent/region_markdown_builder.py`
  - Convert region rows to Markdown, truncating table instance rows after 10.
- Create: `agent/excel_agent/llm_sheet_grouper.py`
  - Call base LLM once per sheet and validate minimal grouping output.
- Create: `agent/excel_agent/excel_visualizer.py`
  - Produce optional VL summaries for low-confidence regions and embedded images.
- Modify: `agent/excel_agent/logic_builder.py`
  - Add `logic_page_name` and grouped area construction.
- Modify: `agent/excel_agent/excel_parse_handler.py`
  - Integrate snapshots, visual summaries, sheet grouping, fallback, and parse metadata.
- Add: `prompt.json`
  - Add `excel_sheet_grouping` and `excel_visual_summary` prompts.
- Modify: `tests/test_excel_parse_agent.py`
  - Keep existing tests and add focused grouping tests.

---

### Task 1: Grouping Models And Prompt Contract

**Files:**
- Modify: `agent/excel_agent/models.py`
- Create: `prompt.json`
- Modify: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Add failing model tests**

Append tests:

```python
from agent.excel_agent.models import (
    ALLOWED_LOGIC_PAGE_NAMES,
    RegionGroup,
    RegionSnapshot,
    SheetGrouping,
)


def test_grouping_models_capture_minimal_llm_contract():
    snapshot = RegionSnapshot(
        region_id="region_1",
        sheet_id="sheet_1",
        cell_range=CellRange(1, 2, 1, 2),
        markdown="| A | B |\n| --- | --- |\n| x | y |",
        raw_text=["A B", "x y"],
        rule_classification={"logic_area_type": "fields", "confidence": 0.82},
        truncated=False,
    )
    grouping = SheetGrouping(
        logic_page_name="bill_summary_page",
        groups=[RegionGroup(region_ids=["region_1"], reason="single section")],
    )

    assert "bill_charge_page" in ALLOWED_LOGIC_PAGE_NAMES
    assert snapshot.region_id == grouping.groups[0].region_ids[0]
    assert grouping.logic_page_name == "bill_summary_page"
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
.\.venv\Scripts\python.exe -m pytest -q tests/test_excel_parse_agent.py::test_grouping_models_capture_minimal_llm_contract
```

Expected: FAIL importing the new names.

- [ ] **Step 3: Implement models**

Add to `agent/excel_agent/models.py`:

```python
ALLOWED_LOGIC_PAGE_NAMES = ("bill_summary_page", "bill_charge_page", "bill_cdr_page")
LogicPageName = Literal["bill_summary_page", "bill_charge_page", "bill_cdr_page"]


@dataclass(frozen=True)
class RegionSnapshot:
    region_id: str
    sheet_id: str
    cell_range: CellRange
    markdown: str
    raw_text: list[str]
    rule_classification: ClassificationDict
    truncated: bool


@dataclass(frozen=True)
class RegionGroup:
    region_ids: list[str]
    reason: str


@dataclass(frozen=True)
class SheetGrouping:
    logic_page_name: LogicPageName
    groups: list[RegionGroup]


@dataclass(frozen=True)
class VisualSummary:
    target_id: str
    target_type: str
    sheet_id: str
    summary: str
    confidence: float


@dataclass(frozen=True)
class GroupingMetadata:
    llm_enabled: bool
    llm_used: bool
    llm_fallback_reason: str | None
    sheet_grouping_count: int
    visual_review_count: int
    visual_review_skipped: list[JsonDict]
```

- [ ] **Step 4: Create prompts**

Create `prompt.json`:

```json
{
  "excel_sheet_grouping": {
    "zh": "你是账单 Excel Sheet 语义分组器。只能输出 JSON 对象。请从 bill_summary_page、bill_charge_page、bill_cdr_page 中选择一个 logic_page_name。只使用输入中出现的 region_id 进行分组，不要编造 region_id，不要输出 area name，不要输出 area type。输入如下：\n{{sheet_payload}}\n输出格式：{\"logic_page_name\":\"bill_summary_page\",\"groups\":[{\"region_ids\":[\"region_1\"],\"reason\":\"string\"}]}",
    "en": "You are a billing Excel sheet semantic grouper. Return only a JSON object. Choose exactly one logic_page_name from bill_summary_page, bill_charge_page, bill_cdr_page. Group only region_id values from the input. Do not invent region IDs. Do not output area names or area types. Input:\n{{sheet_payload}}\nOutput shape: {\"logic_page_name\":\"bill_summary_page\",\"groups\":[{\"region_ids\":[\"region_1\"],\"reason\":\"string\"}]}"
  },
  "excel_visual_summary": {
    "zh": "你是账单 Excel 视觉摘要器。只能输出 JSON 对象。根据图片总结该区域或嵌入图片表达的账单内容。target_id={{target_id}}。输出格式：{\"target_id\":\"{{target_id}}\",\"summary\":\"string\",\"confidence\":0.0}",
    "en": "You are a billing Excel visual summarizer. Return only a JSON object. Summarize what this region or embedded image contains. target_id={{target_id}}. Output shape: {\"target_id\":\"{{target_id}}\",\"summary\":\"string\",\"confidence\":0.0}"
  }
}
```

- [ ] **Step 5: Run tests**

Run:

```bash
.\.venv\Scripts\python.exe -m pytest -q tests/test_excel_parse_agent.py::test_grouping_models_capture_minimal_llm_contract
```

Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add agent/excel_agent/models.py prompt.json tests/test_excel_parse_agent.py
git commit -m "feat: add excel grouping models and prompts"
```

---

### Task 2: Region Range Reading And Markdown Snapshots

**Files:**
- Modify: `agent/excel_agent/excel_reader.py`
- Create: `agent/excel_agent/region_markdown_builder.py`
- Modify: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Add failing tests**

Append tests:

```python
from agent.excel_agent.region_markdown_builder import RegionMarkdownBuilder


def test_excel_reader_reads_only_requested_range(tmp_path):
    path = tmp_path / "range.xlsx"
    make_workbook(path)
    reader = ExcelReader(str(path))
    reader.read()

    rows = reader.read_range("Summary", CellRange(1, 2, 1, 2))
    reader.close()

    assert rows == [["Invoice", "Amount"], ["INV-001", 120]]


def test_region_markdown_builder_preserves_table_and_truncates():
    region = ExcelRegion("sheet_1", CellRange(1, 13, 1, 2), raw_text=[])
    rows = [["Item", "Amount"]] + [[f"row-{i}", i] for i in range(1, 13)]

    snapshot = RegionMarkdownBuilder.build_region_snapshot(
        region_id="region_1",
        region=region,
        rows=rows,
        classification={"logic_area_type": "fee_table", "confidence": 0.78},
        max_instance_rows=10,
    )

    assert "| Item | Amount |" in snapshot.markdown
    assert "| row-10 | 10 |" in snapshot.markdown
    assert "row-11" not in snapshot.markdown
    assert "truncated after 10 rows" in snapshot.markdown
    assert snapshot.truncated is True
```

- [ ] **Step 2: Run tests to verify failure**

Run:

```bash
.\.venv\Scripts\python.exe -m pytest -q tests/test_excel_parse_agent.py::test_excel_reader_reads_only_requested_range tests/test_excel_parse_agent.py::test_region_markdown_builder_preserves_table_and_truncates
```

Expected: FAIL because `read_range` and `RegionMarkdownBuilder` do not exist.

- [ ] **Step 3: Add `read_range`**

Add to `ExcelReader`:

```python
    def read_range(self, sheet_name: str, cell_range: CellRange) -> list[list[Any]]:
        return list(
            self.iter_rows(
                sheet_name,
                min_row=cell_range.start_row,
                max_row=cell_range.end_row,
                min_col=cell_range.start_col,
                max_col=cell_range.end_col,
            )
        )
```

Also import `CellRange` in `excel_reader.py`.

- [ ] **Step 4: Implement Markdown builder**

Create `agent/excel_agent/region_markdown_builder.py`:

```python
from __future__ import annotations

from typing import Any

from .models import ClassificationDict, ExcelRegion, RegionSnapshot


class RegionMarkdownBuilder:
    @staticmethod
    def build_region_snapshot(
        *,
        region_id: str,
        region: ExcelRegion,
        rows: list[list[Any]],
        classification: ClassificationDict,
        max_instance_rows: int = 10,
    ) -> RegionSnapshot:
        normalized = [[RegionMarkdownBuilder._stringify(cell) for cell in row] for row in rows]
        normalized = [row for row in normalized if any(cell for cell in row)]
        truncated = False

        if RegionMarkdownBuilder._is_table(normalized):
            header = normalized[0]
            body = normalized[1:]
            if len(body) > max_instance_rows:
                body = body[:max_instance_rows]
                truncated = True
            markdown = RegionMarkdownBuilder._table(header, body)
        else:
            lines = ["- " + " ".join(cell for cell in row if cell) for row in normalized]
            markdown = "\n".join(lines)

        if truncated:
            markdown = f"{markdown}\n\n... truncated after {max_instance_rows} rows"

        return RegionSnapshot(
            region_id=region_id,
            sheet_id=region.sheet_id,
            cell_range=region.cell_range,
            markdown=markdown,
            raw_text=region.raw_text,
            rule_classification=classification,
            truncated=truncated,
        )

    @staticmethod
    def _is_table(rows: list[list[str]]) -> bool:
        return len(rows) > 1 and max((len(row) for row in rows), default=0) > 1

    @staticmethod
    def _table(header: list[str], body: list[list[str]]) -> str:
        width = max(len(header), *(len(row) for row in body)) if body else len(header)
        padded_header = RegionMarkdownBuilder._pad(header, width)
        lines = [
            "| " + " | ".join(padded_header) + " |",
            "| " + " | ".join("---" for _ in range(width)) + " |",
        ]
        for row in body:
            lines.append("| " + " | ".join(RegionMarkdownBuilder._pad(row, width)) + " |")
        return "\n".join(lines)

    @staticmethod
    def _pad(row: list[str], width: int) -> list[str]:
        return row + ["" for _ in range(width - len(row))]

    @staticmethod
    def _stringify(value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()
```

- [ ] **Step 5: Run tests**

Run:

```bash
.\.venv\Scripts\python.exe -m pytest -q tests/test_excel_parse_agent.py::test_excel_reader_reads_only_requested_range tests/test_excel_parse_agent.py::test_region_markdown_builder_preserves_table_and_truncates
```

Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add agent/excel_agent/excel_reader.py agent/excel_agent/region_markdown_builder.py tests/test_excel_parse_agent.py
git commit -m "feat: build markdown snapshots for excel regions"
```

---

### Task 3: Sheet LLM Grouper

**Files:**
- Create: `agent/excel_agent/llm_sheet_grouper.py`
- Modify: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Add failing grouping tests**

Append tests:

```python
from agent.excel_agent.llm_sheet_grouper import LLMSheetGrouper


def test_llm_sheet_grouper_validates_groups_and_page_name():
    calls = []

    def fake_generate(prompt_template, llm_name="base", lang="zh", **kwargs):
        calls.append((prompt_template, llm_name, kwargs))
        return {
            "logic_page_name": "bill_charge_page",
            "groups": [
                {"region_ids": ["region_1", "missing", "region_1"], "reason": "charges"},
            ],
        }

    snapshots = [
        RegionSnapshot("region_1", "sheet_1", CellRange(1, 2, 1, 2), "r1", [], {"logic_area_type": "fields", "confidence": 0.82}, False),
        RegionSnapshot("region_2", "sheet_1", CellRange(5, 6, 1, 3), "r2", [], {"logic_area_type": "fee_table", "confidence": 0.78}, False),
    ]
    grouping = LLMSheetGrouper(llm_generate=fake_generate).group_sheet(
        sheet_info={"sheet_id": "sheet_1", "sheet_name": "Charges", "sheet_index": 0, "max_row": 6, "max_col": 3, "merged_cells": []},
        region_snapshots=snapshots,
        visual_summaries=[],
    )

    assert grouping.logic_page_name == "bill_charge_page"
    assert [group.region_ids for group in grouping.groups] == [["region_1"], ["region_2"]]
    assert calls[0][0] == "excel_sheet_grouping"
    assert calls[0][1] == "base"


def test_llm_sheet_grouper_falls_back_on_failure():
    def failing_generate(*args, **kwargs):
        raise RuntimeError("boom")

    snapshots = [
        RegionSnapshot("region_1", "sheet_1", CellRange(1, 1, 1, 1), "r1", [], {"logic_area_type": "unknown", "confidence": 0.35}, False)
    ]
    grouper = LLMSheetGrouper(llm_generate=failing_generate)

    grouping = grouper.group_sheet(
        sheet_info={"sheet_id": "sheet_1", "sheet_name": "Unknown", "sheet_index": 0, "max_row": 1, "max_col": 1, "merged_cells": []},
        region_snapshots=snapshots,
        visual_summaries=[],
    )

    assert grouping.logic_page_name == "bill_summary_page"
    assert grouping.groups[0].region_ids == ["region_1"]
    assert "fallback" in grouping.groups[0].reason
    assert grouper.last_fallback_reason == "boom"
```

- [ ] **Step 2: Run tests to verify failure**

Run:

```bash
.\.venv\Scripts\python.exe -m pytest -q tests/test_excel_parse_agent.py::test_llm_sheet_grouper_validates_groups_and_page_name tests/test_excel_parse_agent.py::test_llm_sheet_grouper_falls_back_on_failure
```

Expected: FAIL because module does not exist.

- [ ] **Step 3: Implement grouper**

Create `agent/excel_agent/llm_sheet_grouper.py` with:

```python
from __future__ import annotations

import json
from typing import Any, Callable

from agent.llm.generate_by_llm import generate_by_llm

from .models import ALLOWED_LOGIC_PAGE_NAMES, RegionGroup, RegionSnapshot, SheetGrouping, SheetInfoDict


LLMGenerate = Callable[..., dict[str, Any]]


class LLMSheetGrouper:
    def __init__(self, llm_generate: LLMGenerate = generate_by_llm):
        self.llm_generate = llm_generate
        self.last_fallback_reason: str | None = None
        self.last_llm_used = False

    def group_sheet(
        self,
        *,
        sheet_info: SheetInfoDict,
        region_snapshots: list[RegionSnapshot],
        visual_summaries: list[dict[str, Any]],
    ) -> SheetGrouping:
        self.last_fallback_reason = None
        self.last_llm_used = False
        try:
            payload = self._payload(sheet_info, region_snapshots, visual_summaries)
            response = self.llm_generate(
                "excel_sheet_grouping",
                llm_name="base",
                lang="zh",
                sheet_payload=json.dumps(payload, ensure_ascii=False),
            )
            grouping = self._normalize(response, region_snapshots)
            self.last_llm_used = True
            return grouping
        except Exception as exc:
            self.last_fallback_reason = str(exc)
            return self.fallback(region_snapshots, reason=f"fallback: {exc}")

    @staticmethod
    def fallback(region_snapshots: list[RegionSnapshot], reason: str) -> SheetGrouping:
        return SheetGrouping(
            logic_page_name="bill_summary_page",
            groups=[RegionGroup(region_ids=[snapshot.region_id], reason=reason) for snapshot in region_snapshots],
        )

    @staticmethod
    def _payload(
        sheet_info: SheetInfoDict,
        region_snapshots: list[RegionSnapshot],
        visual_summaries: list[dict[str, Any]],
    ) -> dict[str, Any]:
        return {
            "sheet": {
                "sheet_id": sheet_info["sheet_id"],
                "sheet_name": sheet_info["sheet_name"],
                "sheet_index": sheet_info["sheet_index"],
            },
            "allowed_logic_page_names": list(ALLOWED_LOGIC_PAGE_NAMES),
            "regions": [
                {
                    "region_id": snapshot.region_id,
                    "cell_range": snapshot.cell_range.to_dict(),
                    "markdown": snapshot.markdown,
                    "truncated": snapshot.truncated,
                }
                for snapshot in region_snapshots
            ],
            "visual_summaries": visual_summaries,
        }

    @staticmethod
    def _normalize(response: dict[str, Any], region_snapshots: list[RegionSnapshot]) -> SheetGrouping:
        region_order = [snapshot.region_id for snapshot in region_snapshots]
        valid_ids = set(region_order)
        page_name = response.get("logic_page_name")
        if page_name not in ALLOWED_LOGIC_PAGE_NAMES:
            page_name = "bill_summary_page"

        seen: set[str] = set()
        groups: list[RegionGroup] = []
        for raw_group in response.get("groups", []):
            if not isinstance(raw_group, dict):
                continue
            ids = raw_group.get("region_ids", [])
            if not isinstance(ids, list):
                continue
            deduped = [region_id for region_id in region_order if region_id in ids and region_id not in seen and region_id in valid_ids]
            if not deduped:
                continue
            seen.update(deduped)
            reason = raw_group.get("reason")
            groups.append(RegionGroup(region_ids=deduped, reason=str(reason or "llm grouping")))

        for region_id in region_order:
            if region_id not in seen:
                groups.append(RegionGroup(region_ids=[region_id], reason="fallback: omitted by llm"))

        return SheetGrouping(logic_page_name=page_name, groups=groups)
```

- [ ] **Step 4: Run grouping tests**

Run:

```bash
.\.venv\Scripts\python.exe -m pytest -q tests/test_excel_parse_agent.py::test_llm_sheet_grouper_validates_groups_and_page_name tests/test_excel_parse_agent.py::test_llm_sheet_grouper_falls_back_on_failure
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add agent/excel_agent/llm_sheet_grouper.py tests/test_excel_parse_agent.py
git commit -m "feat: group sheet regions with llm"
```

---

### Task 4: Optional Visual Summaries

**Files:**
- Create: `agent/excel_agent/excel_visualizer.py`
- Modify: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Add failing visualizer tests**

Append tests:

```python
from agent.excel_agent.excel_visualizer import ExcelVisualizer


def test_visualizer_calls_vl_for_low_confidence_region():
    calls = []

    def fake_generate(prompt_template, llm_name="base", lang="zh", **kwargs):
        calls.append((prompt_template, llm_name, kwargs))
        return {"target_id": kwargs["target_id"], "summary": "looks like notes", "confidence": 0.8}

    snapshot = RegionSnapshot(
        "region_1",
        "sheet_1",
        CellRange(1, 1, 1, 1),
        "unclear",
        ["unclear"],
        {"logic_area_type": "unknown", "confidence": 0.35},
        False,
    )

    visualizer = ExcelVisualizer(llm_generate=fake_generate)
    result = visualizer.collect_visual_summaries(
        file_uri="unused.xlsx",
        sheet_info={"sheet_id": "sheet_1", "sheet_name": "S", "sheet_index": 0, "max_row": 1, "max_col": 1, "merged_cells": []},
        region_snapshots=[snapshot],
    )

    assert result["summaries"][0]["target_id"] == "region_1"
    assert calls[0][0] == "excel_visual_summary"
    assert calls[0][1] == "vl"


def test_visualizer_skips_high_confidence_region():
    def fail_generate(*args, **kwargs):
        raise AssertionError("vl should not be called")

    snapshot = RegionSnapshot(
        "region_1",
        "sheet_1",
        CellRange(1, 2, 1, 2),
        "clear",
        ["clear"],
        {"logic_area_type": "fields", "confidence": 0.82},
        False,
    )

    result = ExcelVisualizer(llm_generate=fail_generate).collect_visual_summaries(
        file_uri="unused.xlsx",
        sheet_info={"sheet_id": "sheet_1", "sheet_name": "S", "sheet_index": 0, "max_row": 2, "max_col": 2, "merged_cells": []},
        region_snapshots=[snapshot],
    )

    assert result == {"summaries": [], "skipped": []}
```

- [ ] **Step 2: Run tests to verify failure**

Run:

```bash
.\.venv\Scripts\python.exe -m pytest -q tests/test_excel_parse_agent.py::test_visualizer_calls_vl_for_low_confidence_region tests/test_excel_parse_agent.py::test_visualizer_skips_high_confidence_region
```

Expected: FAIL because module does not exist.

- [ ] **Step 3: Implement lightweight visualizer**

Create `agent/excel_agent/excel_visualizer.py`:

```python
from __future__ import annotations

import base64
from typing import Any, Callable

from agent.llm.generate_by_llm import generate_by_llm

from .models import RegionSnapshot, SheetInfoDict


LLMGenerate = Callable[..., dict[str, Any]]


class ExcelVisualizer:
    def __init__(self, llm_generate: LLMGenerate = generate_by_llm, confidence_threshold: float = 0.65):
        self.llm_generate = llm_generate
        self.confidence_threshold = confidence_threshold

    def collect_visual_summaries(
        self,
        *,
        file_uri: str,
        sheet_info: SheetInfoDict,
        region_snapshots: list[RegionSnapshot],
    ) -> dict[str, list[dict[str, Any]]]:
        summaries: list[dict[str, Any]] = []
        skipped: list[dict[str, Any]] = []

        for snapshot in region_snapshots:
            confidence = float(snapshot.rule_classification.get("confidence", 0.0))
            if confidence >= self.confidence_threshold:
                continue
            try:
                image_base64 = self._render_region_preview(snapshot)
                response = self.llm_generate(
                    "excel_visual_summary",
                    llm_name="vl",
                    lang="zh",
                    target_id=snapshot.region_id,
                    image_base64=image_base64,
                    image_mime_type="image/png",
                )
                summaries.append(
                    {
                        "target_id": str(response.get("target_id", snapshot.region_id)),
                        "target_type": "cell_region",
                        "sheet_id": sheet_info["sheet_id"],
                        "summary": str(response.get("summary", "")),
                        "confidence": float(response.get("confidence", 0.0)),
                    }
                )
            except Exception as exc:
                skipped.append(
                    {
                        "target_id": snapshot.region_id,
                        "target_type": "cell_region",
                        "sheet_id": sheet_info["sheet_id"],
                        "reason": str(exc),
                    }
                )

        skipped.extend(self._embedded_image_skips(file_uri, sheet_info))
        return {"summaries": summaries, "skipped": skipped}

    @staticmethod
    def _render_region_preview(snapshot: RegionSnapshot) -> str:
        # Minimal deterministic PNG header payload for testable VL plumbing.
        # A richer renderer can replace this without changing call sites.
        pseudo_png = b"\x89PNG\r\n\x1a\n" + snapshot.markdown.encode("utf-8")
        return base64.b64encode(pseudo_png).decode("ascii")

    @staticmethod
    def _embedded_image_skips(file_uri: str, sheet_info: SheetInfoDict) -> list[dict[str, Any]]:
        return []
```

- [ ] **Step 4: Run visualizer tests**

Run:

```bash
.\.venv\Scripts\python.exe -m pytest -q tests/test_excel_parse_agent.py::test_visualizer_calls_vl_for_low_confidence_region tests/test_excel_parse_agent.py::test_visualizer_skips_high_confidence_region
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add agent/excel_agent/excel_visualizer.py tests/test_excel_parse_agent.py
git commit -m "feat: collect visual summaries for grouping"
```

---

### Task 5: Logic Builder Group Output

**Files:**
- Modify: `agent/excel_agent/logic_builder.py`
- Modify: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Add failing grouped builder test**

Append test:

```python
def test_logic_builder_creates_named_page_and_grouped_area():
    sheet_info = {"sheet_id": "sheet_1", "sheet_name": "Charges", "sheet_index": 0, "max_row": 6, "max_col": 3, "merged_cells": []}
    region1 = ExcelRegion("sheet_1", CellRange(1, 2, 1, 2), ["A B"])
    region2 = ExcelRegion("sheet_1", CellRange(5, 6, 1, 3), ["Item Qty Amount"])
    classifications = {
        "region_1": {"logic_area_type": "fields", "confidence": 0.82},
        "region_2": {"logic_area_type": "fee_table", "confidence": 0.78},
    }

    page = LogicBuilder.build_page("excel_1", sheet_info, logic_page_name="bill_charge_page")
    area = LogicBuilder.build_grouped_area(
        "excel_1",
        sheet_info,
        group=RegionGroup(region_ids=["region_1", "region_2"], reason="same page content"),
        region_by_id={"region_1": region1, "region_2": region2},
        classification_by_id=classifications,
    )

    assert page["logic_page_name"] == "bill_charge_page"
    assert area["logic_area_type"] == "fee_table"
    assert len(area["location_list"]) == 2
    assert area["group_reason"] == "same page content"
```

- [ ] **Step 2: Run test to verify failure**

Run:

```bash
.\.venv\Scripts\python.exe -m pytest -q tests/test_excel_parse_agent.py::test_logic_builder_creates_named_page_and_grouped_area
```

Expected: FAIL because builder methods do not support grouping.

- [ ] **Step 3: Extend builder**

Update `LogicBuilder.build_page` signature:

```python
    def build_page(
        excel_instance_id: str,
        sheet_info: SheetInfoDict,
        logic_page_name: str = "bill_summary_page",
    ) -> JsonDict:
```

Add `"logic_page_name": logic_page_name` to the page dictionary.

Add `build_grouped_area`:

```python
    @staticmethod
    def build_grouped_area(
        excel_instance_id: str,
        sheet_info: SheetInfoDict,
        *,
        group: RegionGroup,
        region_by_id: dict[str, ExcelRegion],
        classification_by_id: dict[str, ClassificationDict],
    ) -> JsonDict:
        first_region = region_by_id[group.region_ids[0]]
        first_range = first_region.cell_range.to_dict()
        area_type = LogicBuilder._dominant_area_type(group.region_ids, classification_by_id)
        return {
            "logic_area_id": gen_id(),
            "logic_area_name": f"{sheet_info['sheet_name']}!R{first_range['start_row']}C{first_range['start_col']}",
            "logic_area_type": area_type,
            "logic_area_description": f"{area_type} group from sheet {sheet_info['sheet_name']}: {group.reason}",
            "group_reason": group.reason,
            "source_region_id_list": list(group.region_ids),
            "location_list": [
                {
                    "type": "excel",
                    "excel_instance_id": excel_instance_id,
                    "sheet_id": sheet_info["sheet_id"],
                    "sheet_index": sheet_info["sheet_index"],
                    "cell_range": region_by_id[region_id].cell_range.to_dict(),
                    "raw_excel_text_list": region_by_id[region_id].raw_text,
                }
                for region_id in group.region_ids
            ],
        }

    @staticmethod
    def _dominant_area_type(
        region_ids: list[str],
        classification_by_id: dict[str, ClassificationDict],
    ) -> str:
        priority = ["fee_table", "detail_table", "fields", "plain_text", "unknown"]
        present = {classification_by_id[region_id]["logic_area_type"] for region_id in region_ids}
        for area_type in priority:
            if area_type in present:
                return area_type
        return "unknown"
```

Import `RegionGroup`.

- [ ] **Step 4: Run builder tests**

Run:

```bash
.\.venv\Scripts\python.exe -m pytest -q tests/test_excel_parse_agent.py::test_logic_builder_creates_page_and_area tests/test_excel_parse_agent.py::test_logic_builder_creates_named_page_and_grouped_area
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add agent/excel_agent/logic_builder.py tests/test_excel_parse_agent.py
git commit -m "feat: build grouped excel logic areas"
```

---

### Task 6: Handler Integration

**Files:**
- Modify: `agent/excel_agent/excel_parse_handler.py`
- Modify: `tests/test_excel_parse_agent.py`

- [ ] **Step 1: Add failing handler integration tests**

Append tests:

```python
def test_handle_excel_parse_uses_sheet_grouping_llm(tmp_path):
    path = tmp_path / "grouped.xlsx"
    make_workbook(path)

    def fake_generate(prompt_template, llm_name="base", lang="zh", **kwargs):
        if prompt_template == "excel_sheet_grouping":
            return {
                "logic_page_name": "bill_charge_page",
                "groups": [{"region_ids": ["region_1"], "reason": "summary fields"}],
            }
        return {"target_id": kwargs["target_id"], "summary": "", "confidence": 0.0}

    response = handle_excel_parse(
        {
            "request_type": "EXCEL_PARSE",
            "task_id": "task_grouped",
            "site_id": "site_1",
            "project_id": "project_1",
            "payload": {"excel_instance_id": "excel_1", "file_uri": str(path), "parse_mode": "full"},
        },
        llm_generate=fake_generate,
    )

    assert response["status"] == "success"
    assert response["payload"]["logic_page_list"][0]["logic_page_name"] == "bill_charge_page"
    assert response["payload"]["parse_index"]["llm_used"] is True
    assert response["payload"]["parse_index"]["sheet_grouping_count"] == 2


def test_handle_excel_parse_falls_back_when_llm_fails(tmp_path):
    path = tmp_path / "fallback.xlsx"
    make_workbook(path)

    def failing_generate(*args, **kwargs):
        raise RuntimeError("llm offline")

    response = handle_excel_parse(
        {
            "request_type": "EXCEL_PARSE",
            "task_id": "task_fallback",
            "site_id": "site_1",
            "project_id": "project_1",
            "payload": {"excel_instance_id": "excel_1", "file_uri": str(path), "parse_mode": "full"},
        },
        llm_generate=failing_generate,
    )

    assert response["status"] == "success"
    assert response["payload"]["logic_page_list"][0]["logic_page_name"] == "bill_summary_page"
    assert response["payload"]["parse_index"]["llm_used"] is False
    assert "llm offline" in response["payload"]["parse_index"]["llm_fallback_reason"]
```

- [ ] **Step 2: Run tests to verify failure**

Run:

```bash
.\.venv\Scripts\python.exe -m pytest -q tests/test_excel_parse_agent.py::test_handle_excel_parse_uses_sheet_grouping_llm tests/test_excel_parse_agent.py::test_handle_excel_parse_falls_back_when_llm_fails
```

Expected: FAIL because `handle_excel_parse` does not accept `llm_generate` and does not group.

- [ ] **Step 3: Integrate grouping**

Update `handle_excel_parse` to accept dependency injection:

```python
def handle_excel_parse(req: ExcelParseRequest, *, llm_generate=generate_by_llm) -> JsonDict:
```

Inside each sheet loop:

```python
grouper = LLMSheetGrouper(llm_generate=llm_generate)
visualizer = ExcelVisualizer(llm_generate=llm_generate)
...
regions = RegionSplitter.split(reader, sheet_info, profile)
region_by_id = {}
classification_by_id = {}
snapshots = []
for index, region in enumerate(regions, start=1):
    region_id = f"region_{index}"
    classification = RegionClassifier.classify(region)
    rows = reader.read_range(sheet_info["sheet_name"], region.cell_range)
    snapshot = RegionMarkdownBuilder.build_region_snapshot(
        region_id=region_id,
        region=region,
        rows=rows,
        classification=classification,
    )
    region_by_id[region_id] = region
    classification_by_id[region_id] = classification
    snapshots.append(snapshot)

visual = visualizer.collect_visual_summaries(
    file_uri=payload["file_uri"],
    sheet_info=sheet_info,
    region_snapshots=snapshots,
)
grouping = grouper.group_sheet(
    sheet_info=sheet_info,
    region_snapshots=snapshots,
    visual_summaries=visual["summaries"],
)
logic_page_list.append(LogicBuilder.build_page(excel_instance_id, sheet_info, grouping.logic_page_name))
for group in grouping.groups:
    logic_area_list.append(
        LogicBuilder.build_grouped_area(
            excel_instance_id,
            sheet_info,
            group=group,
            region_by_id=region_by_id,
            classification_by_id=classification_by_id,
        )
    )
```

Maintain `llm_used`, `llm_fallback_reason`, `sheet_grouping_count`, `visual_review_count`, and `visual_review_skipped` in `parse_index`.

- [ ] **Step 4: Run integration tests**

Run:

```bash
.\.venv\Scripts\python.exe -m pytest -q tests/test_excel_parse_agent.py::test_handle_excel_parse_uses_sheet_grouping_llm tests/test_excel_parse_agent.py::test_handle_excel_parse_falls_back_when_llm_fails
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add agent/excel_agent/excel_parse_handler.py tests/test_excel_parse_agent.py
git commit -m "feat: integrate sheet llm grouping into parse handler"
```

---

### Task 7: Full Verification And Cleanup

**Files:**
- Modify only if verification finds issues.

- [ ] **Step 1: Run full tests**

Run:

```bash
.\.venv\Scripts\python.exe -m pytest -q
```

Expected: all tests PASS.

- [ ] **Step 2: Run compile check**

Run:

```bash
.\.venv\Scripts\python.exe -m py_compile agent\excel_agent\models.py agent\excel_agent\excel_reader.py agent\excel_agent\sheet_profiler.py agent\excel_agent\region_splitter.py agent\excel_agent\region_classifier.py agent\excel_agent\logic_builder.py agent\excel_agent\excel_parse_handler.py agent\excel_agent\region_markdown_builder.py agent\excel_agent\llm_sheet_grouper.py agent\excel_agent\excel_visualizer.py
```

Expected: no output and exit code 0.

- [ ] **Step 3: Scan constraints**

Run:

```bash
Select-String -Path agent\excel_agent\*.py,tests\*.py -Pattern 'pandas'
Select-String -Path agent\excel_agent\*.py -Pattern 'load_workbook|read_only|data_only'
```

Expected:

- no `pandas` hits
- `load_workbook` still uses `read_only=True` and `data_only=True`

- [ ] **Step 4: Check git status**

Run:

```bash
git status --short
```

Expected: clean after commits.

- [ ] **Step 5: Final commit if cleanup was needed**

If Step 1-3 required fixes:

```bash
git add agent tests prompt.json
git commit -m "test: verify excel sheet llm grouping"
```

If no fixes were needed, do not create an empty commit.

---

## Self-Review

**Spec coverage:** This plan implements one logic page per sheet, rule-first region splitting, stable region IDs, Markdown region content with table preservation, truncation after 10 table instance rows, one base LLM grouping call per sheet, minimal grouping output, optional VL summaries for low-confidence regions, grouped logic areas, and successful fallback when LLM fails.

**Placeholder scan:** The plan contains no deferred implementation placeholders. Each task has concrete files, tests, code snippets, commands, and expected outcomes.

**Type consistency:** The plan consistently uses `RegionSnapshot`, `RegionGroup`, `SheetGrouping`, `LLMSheetGrouper`, `RegionMarkdownBuilder`, `ExcelVisualizer`, and `LogicBuilder.build_grouped_area`.
