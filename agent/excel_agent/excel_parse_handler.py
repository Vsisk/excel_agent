from __future__ import annotations

from typing import Any, Callable

from .excel_reader import ExcelReader
from .grouping_memory import EmbeddingGenerate, WorkbookGroupingMemory
from .llm_sheet_grouper import LLMSheetGrouper
from .models import ClassificationDict, ExcelParseRequest, JsonDict, RegionGroup, RegionSnapshot, SheetInfoDict
from .region_markdown_builder import RegionMarkdownBuilder
from .region_splitter import RegionSplitter
from .sheet_profiler import SheetProfiler


LLMGenerate = Callable[..., dict[str, Any]]
STRUCTURAL_CLASSIFICATION: ClassificationDict = {"logic_area_type": "unknown", "confidence": 1.0}


def handle_excel_parse(
    req: ExcelParseRequest,
    *,
    llm_generate: LLMGenerate | None = None,
    embedding_generate: EmbeddingGenerate | None = None,
) -> list[dict[str, Any]]:
    _validate_request(req)
    payload = req["payload"]
    reader = ExcelReader(payload["file_uri"])
    grouper = LLMSheetGrouper(llm_generate=llm_generate) if llm_generate else LLMSheetGrouper()
    grouping_memory = WorkbookGroupingMemory(embedding_generate=embedding_generate)

    sheet_content: list[JsonDict] = []
    sheet_profiles: dict[str, JsonDict] = {}
    llm_used = False
    sheet_grouping_count = 0

    try:
        workbook = reader.read()
        for sheet_info in workbook["sheet_list"]:
            profile = SheetProfiler.profile(reader, sheet_info)
            sheet_profiles[sheet_info["sheet_id"]] = profile.to_dict()
            regions = RegionSplitter.split(reader, sheet_info, profile)
            region_by_id = {}
            snapshots = []
            snapshot_by_id = {}

            for index, region in enumerate(regions, start=1):
                region_id = f"region_{index}"
                rows = reader.read_range(sheet_info["sheet_name"], region.cell_range)
                snapshot = RegionMarkdownBuilder.build_region_snapshot(
                    region_id=region_id,
                    region=region,
                    rows=rows
                )
                region_by_id[region_id] = region
                snapshots.append(snapshot)
                snapshot_by_id[region_id] = snapshot

            visual_summaries: list[JsonDict] = []
            memory_matches = grouping_memory.retrieve_for_sheet(
                sheet_info=sheet_info,
                region_snapshots=snapshots,
                visual_summaries=visual_summaries,
            )
            grouping_memory.record_sheet_matches(sheet_info, memory_matches)
            grouping = grouper.group_sheet(
                sheet_info=sheet_info,
                region_snapshots=snapshots,
                grouping_memory_matches=memory_matches,
            )
            sheet_grouping_count += 1
            llm_used = llm_used or grouper.last_llm_used
            grouping_memory.record_consistency_warnings(sheet_info, memory_matches, grouping)
            grouping_memory.remember_sheet_grouping(
                sheet_info=sheet_info,
                region_snapshots=snapshots,
                visual_summaries=visual_summaries,
                grouping=grouping,
            )

            sheet_content.append(
                _build_sheet_content(
                    sheet_info=sheet_info,
                    page_type=grouping.logic_page_name,
                    groups=grouping.groups,
                    snapshot_by_id=snapshot_by_id,
                )
            )
        return sheet_content

    finally:
        reader.close()




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


def _build_sheet_content(
    *,
    sheet_info: SheetInfoDict,
    page_type: str,
    groups: list[RegionGroup],
    snapshot_by_id: dict[str, RegionSnapshot],
) -> JsonDict:
    return {
        "page_id": sheet_info["sheet_index"] + 1,
        "page_type": page_type,
        "blocks": [
            _build_group_block(group_index=index, group=group, snapshot_by_id=snapshot_by_id)
            for index, group in enumerate(groups, start=1)
        ],
    }


def _build_group_block(
    *,
    group_index: int,
    group: RegionGroup,
    snapshot_by_id: dict[str, RegionSnapshot],
) -> JsonDict:
    snapshots = [snapshot_by_id[region_id] for region_id in group.region_ids if region_id in snapshot_by_id]
    return {
        "group_id": f"group_{group_index}",
        "bbox": _build_group_bbox(snapshots),
        "table_md": _combine_table_markdown(snapshots),
    }


def _build_group_bbox(snapshots: list[RegionSnapshot]) -> JsonDict:
    if not snapshots:
        return {"left": 0, "right": 0, "top": 0, "bottom": 0}

    return {
        "left": min(snapshot.cell_range.start_col for snapshot in snapshots),
        "right": max(snapshot.cell_range.end_col for snapshot in snapshots),
        "top": min(snapshot.cell_range.start_row for snapshot in snapshots),
        "bottom": max(snapshot.cell_range.end_row for snapshot in snapshots),
    }


def _combine_table_markdown(snapshots: list[RegionSnapshot]) -> str:
    lines: list[str] = []
    for snapshot in snapshots:
        for line in snapshot.markdown.splitlines():
            normalized = line.strip()
            if not normalized or _is_markdown_separator_row(normalized):
                continue
            lines.append(normalized)
    return "\n".join(lines)


def _is_markdown_separator_row(line: str) -> bool:
    cells = [cell.strip() for cell in line.strip("|").split("|")]
    return bool(cells) and all(cell and set(cell) <= {"-", ":"} for cell in cells)
