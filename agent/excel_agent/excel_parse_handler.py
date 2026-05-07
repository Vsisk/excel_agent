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
