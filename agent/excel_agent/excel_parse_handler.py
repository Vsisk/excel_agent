from __future__ import annotations

from typing import Any, Callable

from .excel_visualizer import ExcelVisualizer
from .excel_reader import ExcelReader
from .llm_sheet_grouper import LLMSheetGrouper
from .logic_builder import LogicBuilder
from .models import ClassificationDict, ExcelParseRequest, JsonDict
from .region_markdown_builder import RegionMarkdownBuilder
from .region_splitter import RegionSplitter
from .sheet_profiler import SheetProfiler


LLMGenerate = Callable[..., dict[str, Any]]
STRUCTURAL_CLASSIFICATION: ClassificationDict = {"logic_area_type": "unknown", "confidence": 1.0}


def handle_excel_parse(req: ExcelParseRequest, *, llm_generate: LLMGenerate | None = None) -> JsonDict:
    _validate_request(req)
    payload = req["payload"]
    excel_instance_id = payload["excel_instance_id"]
    reader = ExcelReader(payload["file_uri"])
    grouper = LLMSheetGrouper(llm_generate=llm_generate) if llm_generate else LLMSheetGrouper()
    visualizer = ExcelVisualizer(llm_generate=llm_generate) if llm_generate else ExcelVisualizer()

    logic_page_list: list[JsonDict] = []
    logic_area_list: list[JsonDict] = []
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

            for index, region in enumerate(regions, start=1):
                region_id = f"region_{index}"
                classification = STRUCTURAL_CLASSIFICATION
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
            visual_review_count += len(visual["summaries"])
            visual_review_skipped.extend(visual["skipped"])

            grouping = grouper.group_sheet(
                sheet_info=sheet_info,
                region_snapshots=snapshots,
                visual_summaries=visual["summaries"],
            )
            sheet_grouping_count += 1
            llm_used = llm_used or grouper.last_llm_used
            if grouper.last_fallback_reason:
                llm_fallback_reasons.append(grouper.last_fallback_reason)

            logic_page_list.append(
                LogicBuilder.build_page(
                    excel_instance_id,
                    sheet_info,
                    logic_page_name=grouping.logic_page_name,
                )
            )
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
                "llm_enabled": True,
                "llm_used": llm_used,
                "llm_fallback_reason": "; ".join(llm_fallback_reasons) if llm_fallback_reasons else None,
                "sheet_grouping_count": sheet_grouping_count,
                "visual_review_count": visual_review_count,
                "visual_review_skipped": visual_review_skipped,
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
