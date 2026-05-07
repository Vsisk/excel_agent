# Excel Sheet LLM Grouping Design

## Goal

Add an optional sheet-level LLM grouping stage to the Excel parse agent. For each sheet, the agent first uses rules to split regions, converts each region to compact Markdown, then calls LLM once to choose the sheet's logic page name and group related regions.

The LLM grouping output is intentionally minimal so later flows can reuse the same downstream logic for both whole-sheet parsing and single-region input.

## Corrected Scope

The LLM does not own low-level parsing, region splitting, area naming, area typing, or cross-sheet binding in this stage.

It only returns:

```json
{
  "logic_page_name": "bill_summary_page",
  "groups": [
    {
      "region_ids": ["region_1", "region_2"],
      "reason": "string"
    }
  ]
}
```

Valid `logic_page_name` values are:

- `bill_summary_page`
- `bill_charge_page`
- `bill_cdr_page`

Each Excel sheet still maps to exactly one `logic_page`.

## Current Context

The current pipeline is:

1. `ExcelReader` opens workbooks with `openpyxl` using `read_only=True` and `data_only=True`.
2. `SheetProfiler` computes used ranges and row/column densities.
3. `RegionSplitter` creates initial rectangular regions from rule-based density signals.
4. `RegionClassifier` assigns a rule-based type and confidence.
5. `LogicBuilder` emits `logic_page_list` and `logic_area_list`.
6. `excel_parse_handler.handle_excel_parse()` returns `EXCEL_PARSE_RESULT`.

The new `agent/llm/` package provides:

- `generate_by_llm(prompt_template, llm_name="base" | "vl", image_base64=...)`
- `LLMClient`
- JSON response parsing
- `.env` based OpenAI-compatible settings

There is not yet a `prompt.json`.

## Components

### `models.py`

Add internal structures for stable region grouping:

- `RegionSnapshot`
  - `region_id`
  - `sheet_id`
  - `cell_range`
  - `markdown`
  - `raw_text`
  - `rule_classification`
  - `truncated`
- `SheetGrouping`
  - `logic_page_name`
  - `groups`
- `RegionGroup`
  - `region_ids`
  - `reason`
- `GroupingResult`
  - per-sheet groupings
  - fallback metadata
  - visual review metadata

These are internal agent structures. The external WS request shape stays unchanged.

### `region_markdown_builder.py`

Responsibilities:

- Convert each rule-split region to Markdown before LLM grouping.
- Preserve table structure where possible.
- Truncate large table instances.

Public interface:

```python
class RegionMarkdownBuilder:
    @staticmethod
    def build_region_snapshot(
        *,
        region_id: str,
        region: ExcelRegion,
        rows: list[list[object]],
        classification: ClassificationDict,
        max_instance_rows: int = 10,
    ) -> RegionSnapshot:
        ...
```

Markdown rules:

- If the region has more than one row and more than one column, render as a Markdown table.
- If the region looks like key-value rows, render as a two-column Markdown table.
- Otherwise render as bullet lines or plain text.
- For data tables with more than 10 instance rows, keep the header and first 10 instance rows.
- Mark truncated snapshots with `truncated=true` and include a note in Markdown such as `... truncated after 10 rows`.

The builder may read only the region range, not the whole sheet.

### `excel_visualizer.py`

Responsibilities:

- Generate optional visual review inputs for:
  - low-confidence rule regions
  - embedded workbook images
- Produce VL summaries that can be included as extra context in the sheet-level grouping prompt.

The visual stage supplements grouping context. It does not directly create final groups or logic areas.

Low-confidence cell regions:

- A region is low-confidence when rule classification confidence is below `0.65`.
- The visualizer renders a lightweight PNG preview of the cell range from values and cell coordinates.
- VL returns a compact summary of what the region appears to contain.

Embedded images:

- The visualizer detects images with `openpyxl` where possible.
- Each embedded image is sent to VL for a compact summary.
- If anchor metadata can identify a sheet and approximate cell location, that metadata is included in grouping context.

If rendering or extraction fails, the visualizer records a skip reason. Failure does not fail `EXCEL_PARSE`.

### `llm_sheet_grouper.py`

Responsibilities:

- For each sheet, call base LLM once with:
  - sheet metadata
  - region IDs
  - region ranges
  - region Markdown
  - truncation flags
  - optional VL summaries
- Validate the LLM output.
- Return one grouping result per sheet.

Public interface:

```python
class LLMSheetGrouper:
    def group_sheet(
        self,
        *,
        sheet_info: SheetInfoDict,
        region_snapshots: list[RegionSnapshot],
        visual_summaries: list[dict],
    ) -> SheetGrouping:
        ...
```

Validation rules:

- `logic_page_name` must be one of the three allowed values.
- Every `region_id` in every group must exist in the current sheet.
- A region can appear in at most one group.
- Empty groups are ignored.
- Regions omitted by LLM are added as single-region fallback groups.
- Duplicate region IDs inside one group are deduplicated in original region order.
- Invalid LLM responses trigger sheet-level fallback grouping.

Fallback grouping:

```json
{
  "logic_page_name": "bill_summary_page",
  "groups": [
    {
      "region_ids": ["region_1"],
      "reason": "fallback: llm unavailable"
    }
  ]
}
```

### `prompt.json`

Add these prompt keys:

- `excel_sheet_grouping`
- `excel_visual_summary`

`excel_sheet_grouping` requires strict JSON object output:

```json
{
  "logic_page_name": "bill_summary_page | bill_charge_page | bill_cdr_page",
  "groups": [
    {
      "region_ids": ["region_1", "region_2"],
      "reason": "string"
    }
  ]
}
```

The prompt must tell the model:

- Choose exactly one page name from the allowed list.
- Group only region IDs from the input.
- Do not invent region IDs.
- Do not output area names or area types.
- Keep groups semantically cohesive.

`excel_visual_summary` returns:

```json
{
  "target_id": "string",
  "summary": "string",
  "confidence": 0.0
}
```

## Data Flow

For each `EXCEL_PARSE` request:

1. Validate the WS request.
2. Read workbook metadata.
3. For each sheet:
   - profile sheet
   - split regions by rules
   - rule-classify regions
   - assign stable region IDs
   - read each region range and build Markdown snapshot
   - collect low-confidence visual summaries
   - collect embedded image visual summaries
   - call base LLM once for sheet grouping
   - fallback to one-region groups if LLM fails
4. Build one `logic_page` per sheet using the returned `logic_page_name`.
5. Build one `logic_area` per returned group.
6. Return `EXCEL_PARSE_RESULT`.

## Logic Output Mapping

### `logic_page`

Add `logic_page_name` while preserving the existing relation shape:

```json
{
  "logic_page_id": "string",
  "logic_page_name": "bill_summary_page",
  "logic_page_relation": {
    "type": "excel",
    "excel_instance_id": "string",
    "sheet_id": "string",
    "sheet_name": "string",
    "sheet_index": 0
  }
}
```

### `logic_area`

Each group becomes one logic area.

Because the grouping prompt no longer outputs area name or area type:

- `logic_area_name` is generated locally from sheet name and first region range.
- `logic_area_type` is derived locally from the grouped regions' rule classifications.
- `logic_area_description` is generated locally and can include the group reason.
- `location_list` contains one location per grouped region.

The group reason can be recorded as internal metadata, for example:

```json
{
  "group_reason": "same charge table split by blank rows"
}
```

## LLM Failure Behavior

LLM is optional and best-effort.

If base LLM, VL LLM, visual rendering, or image extraction fails:

- `status` remains `"success"`
- rule regions are preserved
- each region becomes its own group unless a valid grouping exists
- `logic_page_name` defaults to `bill_summary_page`
- `parse_index` records:
  - `llm_enabled`
  - `llm_used`
  - `llm_fallback_reason`
  - `sheet_grouping_count`
  - `visual_review_count`
  - `visual_review_skipped`

## Single-Region Compatibility

The grouping output intentionally excludes fields that belong to later stages:

- no area type
- no area name
- no final schema mapping
- no EDSL output

This allows a future single-region flow to reuse the same downstream area naming/type-generation step without depending on sheet-level grouping prompts.

## Testing Strategy

Tests should not require real LLM calls.

Use fake LLM functions to cover:

- sheet grouping returns valid page name and groups
- invalid page name falls back to `bill_summary_page`
- invented region IDs are ignored
- omitted region IDs are added as single-region fallback groups
- duplicate region IDs are deduplicated
- table Markdown preserves header and rows
- table Markdown truncates after 10 instance rows
- low-confidence region triggers VL summary
- embedded image target triggers VL summary when present
- LLM unavailable returns success with fallback metadata
- no low-confidence region and no embedded image avoids VL calls

Keep the existing rule parser tests.

## Non-Goals

This change does not add:

- frontend logic
- HTTP endpoints
- EDSL writing
- real LLM dependency in tests
- GUI Excel automation
- LLM-generated area names
- LLM-generated area types
- cross-sheet logic page binding

## Self-Review

The design matches the corrected requirement: each sheet maps to one logic page, rules split regions first, one base LLM call per sheet groups existing region IDs and chooses one of three logic page names, Markdown preserves table structure with truncation after 10 instance rows, and LLM failure degrades to successful rule output. It keeps the grouping prompt minimal for future single-region compatibility.
