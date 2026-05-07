# Excel LLM Refinement Design

## Goal

Add an optional LLM refinement stage to the Excel parse agent so `EXCEL_PARSE` can improve region classification, region merge/split decisions, logic page binding, and visual review while preserving the current rule-based parser as a reliable fallback.

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

## Design Choice

Use an additive LLM refinement layer after initial rule parsing and before logic output building.

The rule parser remains the source of safe fallback behavior. LLM output is treated as advisory and must be normalized and validated before it changes regions or classifications.

If LLM is disabled, unusable, or raises an error, the handler still returns:

```json
{
  "request_type": "EXCEL_PARSE_RESULT",
  "status": "success"
}
```

with the rule-based result. `parse_index` records the fallback reason.

## Components

### `excel_visualizer.py`

Responsibilities:

- Detect embedded images in Excel sheets.
- Create visual review payloads for:
  - low-confidence cell regions
  - embedded workbook images
- Return base64 PNG image payloads plus metadata that identifies the source.

Expected public interface:

```python
class ExcelVisualizer:
    @staticmethod
    def collect_visual_targets(
        file_uri: str,
        sheet_info_list: list[SheetInfoDict],
        classified_regions: list[ClassifiedRegion],
        confidence_threshold: float,
    ) -> VisualCollection:
        ...
```

Visual target types:

- `cell_region`
- `embedded_image`

For `cell_region`, the metadata includes `sheet_id`, `sheet_name`, `sheet_index`, and `cell_range`.

For `embedded_image`, the metadata includes `sheet_id`, `sheet_name`, `sheet_index`, image index, and anchor details when available.

Implementation preference:

- Use `openpyxl` for embedded image extraction where possible.
- For low-confidence cell range previews, generate a deterministic lightweight PNG rendering from cell values rather than automating Excel. This keeps tests reliable and avoids GUI requirements.
- If a visual target cannot be generated, skip it and record a structured skip reason.

### `llm_region_refiner.py`

Responsibilities:

- Build compact region summaries for base LLM review.
- Call base LLM for:
  - classification correction
  - merge suggestions
  - split suggestions
  - logic page binding
- Call VL LLM for:
  - low-confidence cell region screenshots
  - embedded image review
- Normalize and validate all LLM suggestions.
- Apply valid suggestions to produce refined classified regions and page binding metadata.
- Return fallback metadata when LLM is unavailable or fails.

Expected public interface:

```python
class LLMRegionRefiner:
    def refine(
        self,
        *,
        excel_instance_id: str,
        workbook: ExcelWorkbookDict,
        regions_by_sheet: dict[str, list[ExcelRegion]],
        classifications_by_region: dict[str, ClassificationDict],
        file_uri: str,
    ) -> RefinementResult:
        ...
```

The refiner accepts dependency injection for tests:

```python
LLMRegionRefiner(
    llm_generate=generate_by_llm,
    visualizer=ExcelVisualizer(),
    confidence_threshold=0.65,
)
```

### `models.py`

Add small internal dataclasses or `TypedDict`s for the refinement stage:

- `ClassifiedRegion`
- `VisualTarget`
- `VisualCollection`
- `LLMRegionAction`
- `PageBinding`
- `RefinementResult`

These models are internal to the agent. The external WS input/output protocol does not change.

### `prompt.json`

Add prompt keys:

- `excel_region_refine`
- `excel_region_visual_review`

Both prompts require strict JSON object output.

`excel_region_refine` returns:

```json
{
  "region_updates": [
    {
      "region_id": "string",
      "logic_area_type": "fields | fee_table | detail_table | plain_text | unknown",
      "confidence": 0.0,
      "reason": "string"
    }
  ],
  "merge_suggestions": [
    {
      "source_region_ids": ["string"],
      "logic_area_type": "fields | fee_table | detail_table | plain_text | unknown",
      "reason": "string"
    }
  ],
  "split_suggestions": [
    {
      "source_region_id": "string",
      "cell_ranges": [
        {
          "start_row": 1,
          "end_row": 1,
          "start_col": 1,
          "end_col": 1
        }
      ],
      "reason": "string"
    }
  ],
  "page_bindings": [
    {
      "region_id": "string",
      "sheet_id": "string",
      "reason": "string"
    }
  ]
}
```

`excel_region_visual_review` returns:

```json
{
  "target_id": "string",
  "logic_area_type": "fields | fee_table | detail_table | plain_text | unknown",
  "confidence": 0.0,
  "recommended_action": "keep | update | split | merge | ignore",
  "reason": "string"
}
```

## Data Flow

The enhanced `EXCEL_PARSE` flow is:

1. Validate request.
2. Read workbook metadata.
3. For each sheet:
   - profile sheet
   - split regions
   - rule-classify regions
4. Build `ClassifiedRegion` records with stable internal region IDs.
5. Call `LLMRegionRefiner.refine()`.
6. If refinement succeeds:
   - apply valid classification updates
   - apply safe merges
   - apply safe splits
   - apply page bindings
7. If refinement fails:
   - keep rule result
   - record fallback metadata
8. Build `logic_page_list` and `logic_area_list`.
9. Return `EXCEL_PARSE_RESULT`.

## Merge And Split Rules

LLM suggestions are accepted only when they are safe:

- All referenced region IDs must exist.
- Merge candidates must be on the same sheet.
- Merge output must remain a rectangular bounding range.
- Split ranges must be inside the source region.
- Split ranges must use 1-based indexes.
- Split ranges must not overlap.
- Invalid or partial suggestions are ignored, not fatal.

When merging regions, `raw_text` is concatenated in row-major region order.

When splitting regions, raw text is regenerated from the workbook for the suggested sub-ranges.

## Logic Page Binding

Each `logic_page` still maps to one Excel sheet.

The LLM can suggest a page binding for each region. The first implementation only accepts bindings that point to the region's own sheet. This keeps the output stable while leaving room for future cross-sheet semantic grouping.

Accepted binding metadata is recorded in `parse_index`, but the external `logic_page_relation` shape stays unchanged.

## Visual Review

VL review is triggered for:

- rule or base-LLM classified regions with confidence below `0.65`
- embedded images found in the workbook

Visual review can update a region classification when:

- the target maps to an existing region, and
- VL confidence is greater than the current confidence.

Embedded images that do not map to an existing region produce review metadata in `parse_index`. They do not create a `logic_area` unless the image can be associated with a sheet anchor and a valid cell range.

## Error Handling And Fallback

LLM refinement is best-effort.

On any LLM or visualizer error:

- `status` remains `"success"`
- rule-based logic output is returned
- `parse_index` includes:
  - `llm_enabled`
  - `llm_used`
  - `llm_fallback_reason`
  - `visual_review_count`
  - `visual_review_skipped`

Example:

```json
{
  "llm_enabled": true,
  "llm_used": false,
  "llm_fallback_reason": "LLM settings are not usable",
  "visual_review_count": 0,
  "visual_review_skipped": [
    {
      "target_type": "cell_region",
      "reason": "visual target generation unavailable"
    }
  ]
}
```

## Testing Strategy

Tests should not require real LLM calls.

Use fake LLM functions to cover:

- base LLM classification update
- safe merge suggestion
- safe split suggestion
- invalid suggestion ignored
- low-confidence region triggers VL review
- embedded image target triggers VL review when present
- LLM unavailable or failing returns rule output with fallback metadata
- no low-confidence region and no embedded image avoids VL calls

Keep the current rule-parser tests.

Add focused tests for the new refiner and handler integration.

## Non-Goals

This change does not add:

- frontend logic
- HTTP endpoints
- EDSL writing
- hard dependency on real LLM availability
- GUI Excel automation
- cross-sheet semantic page merging

## Self-Review

The design keeps the existing WS protocol unchanged and adds only optional enrichment. It covers region classification, merging, splitting, logic page binding, low-confidence region screenshots, embedded image review, and fallback behavior. It avoids placeholders and makes invalid LLM output non-fatal.
