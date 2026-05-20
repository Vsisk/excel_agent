from __future__ import annotations

from typing import Any, Callable

from .excel_visualizer import ExcelVisualizer
from .excel_reader import ExcelReader
from .grouping_memory import EmbeddingGenerate, WorkbookGroupingMemory
from .llm_sheet_grouper import LLMSheetGrouper
from .logic_builder import LogicBuilder
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
) -> JsonDict:
    _validate_request(req)
    payload = req["payload"]
    excel_instance_id = payload["excel_instance_id"]
    reader = ExcelReader(payload["file_uri"])
    grouper = LLMSheetGrouper(llm_generate=llm_generate) if llm_generate else LLMSheetGrouper()
    visualizer = ExcelVisualizer(llm_generate=llm_generate) if llm_generate else ExcelVisualizer()
    grouping_memory = WorkbookGroupingMemory(embedding_generate=embedding_generate)

    logic_page_list: list[JsonDict] = []
    logic_area_list: list[JsonDict] = []
    sheet_content: list[JsonDict] = []
    sheet_profiles: dict[str, JsonDict] = {}
    llm_used = False
    llm_fallback_reasons: list[str] = []
    sheet_grouping_count = 0
    visual_review_count = 0
    visual_review_skipped: list[JsonDict] = []

    try:
        workbook = reader.read()
        for sheet_info in workbook["sheet_list"]:
            profile = SheetProfiler.profile(reader, sheet_info)
            sheet_profiles[sheet_info["sheet_id"]] = profile.to_dict()
            regions = RegionSplitter.split(reader, sheet_info, profile)
            region_by_id = {}
            classification_by_id = {}
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

            visual = visualizer.collect_visual_summaries(
                file_uri=payload["file_uri"],
                sheet_info=sheet_info,
                region_snapshots=snapshots,
            )
            visual_review_count += len(visual["summaries"])
            visual_review_skipped.extend(visual["skipped"])

            memory_matches = grouping_memory.retrieve_for_sheet(
                sheet_info=sheet_info,
                region_snapshots=snapshots,
                visual_summaries=visual["summaries"],
            )
            grouping_memory.record_sheet_matches(sheet_info, memory_matches)
            grouping = grouper.group_sheet(
                sheet_info=sheet_info,
                region_snapshots=snapshots,
                visual_summaries=visual["summaries"],
                grouping_memory_matches=memory_matches,
            )
            sheet_grouping_count += 1
            llm_used = llm_used or grouper.last_llm_used
            if grouper.last_fallback_reason:
                llm_fallback_reasons.append(grouper.last_fallback_reason)
            grouping_memory.record_consistency_warnings(sheet_info, memory_matches, grouping)
            grouping_memory.remember_sheet_grouping(
                sheet_info=sheet_info,
                region_snapshots=snapshots,
                visual_summaries=visual["summaries"],
                grouping=grouping,
            )

            logic_page_list.append(
                LogicBuilder.build_page(
                    excel_instance_id,
                    sheet_info,
                    logic_page_name=grouping.logic_page_name,
                )
            )
            sheet_content.append(
                _build_sheet_content(
                    sheet_info=sheet_info,
                    page_type=grouping.logic_page_name,
                    groups=grouping.groups,
                    snapshot_by_id=snapshot_by_id,
                )
            )
            for group in grouping.groups:
                logic_area_list.append(
                    LogicBuilder.build_grouped_area(
                        excel_instance_id,
                        sheet_info,
                        group=group,
                        region_by_id=region_by_id
                    )
                )
    finally:
        reader.close()

    return {
        "request_type": "EXCEL_PARSE_RESULT",
        "task_id": req["task_id"],
        "status": "success",
        "payload": {
            "sheet_content": sheet_content,
            "logic_page_list": logic_page_list,
            "logic_area_list": logic_area_list,
            "parse_index": {
                "excel_instance_id": excel_instance_id,
                "sheet_count": len(logic_page_list),
                "area_count": len(logic_area_list),
                "sheet_profiles": sheet_profiles,
                "llm_enabled": True,
                "llm_used": llm_used,
                "llm_fallback_reason": "; ".join(llm_fallback_reasons) if llm_fallback_reasons else None,
                "sheet_grouping_count": sheet_grouping_count,
                "visual_review_count": visual_review_count,
                "visual_review_skipped": visual_review_skipped,
                "grouping_memory_enabled": grouping_memory.enabled,
                "grouping_memory_used": grouping_memory.used,
                "grouping_memory_fallback_reason": grouping_memory.fallback_reason,
                "grouping_memory_template_count": grouping_memory.template_count,
                "grouping_memory_matches": grouping_memory.matches_by_sheet,
                "memory_consistency_warnings": grouping_memory.warnings,
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
