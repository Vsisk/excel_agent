import re
from pathlib import Path

import pytest
from openpyxl import Workbook

from agent.excel_agent.excel_parse_handler import handle_excel_parse
from agent.excel_agent.excel_reader import ExcelReader
from agent.excel_agent.llm_sheet_grouper import LLMSheetGrouper
from agent.excel_agent.logic_builder import LogicBuilder
from agent.excel_agent.models import (
    ALLOWED_LOGIC_PAGE_NAMES,
    CellRange,
    ExcelRegion,
    RegionGroup,
    RegionSnapshot,
    SheetGrouping,
    gen_id,
)
from agent.excel_agent.region_classifier import RegionClassifier
from agent.excel_agent.region_markdown_builder import RegionMarkdownBuilder
from agent.excel_agent.region_splitter import RegionSplitter
from agent.excel_agent.sheet_profiler import SheetProfiler


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
        RegionSnapshot(
            "region_1",
            "sheet_1",
            CellRange(1, 2, 1, 2),
            "r1",
            [],
            {"logic_area_type": "fields", "confidence": 0.82},
            False,
        ),
        RegionSnapshot(
            "region_2",
            "sheet_1",
            CellRange(5, 6, 1, 3),
            "r2",
            [],
            {"logic_area_type": "fee_table", "confidence": 0.78},
            False,
        ),
    ]
    grouping = LLMSheetGrouper(llm_generate=fake_generate).group_sheet(
        sheet_info={
            "sheet_id": "sheet_1",
            "sheet_name": "Charges",
            "sheet_index": 0,
            "max_row": 6,
            "max_col": 3,
            "merged_cells": [],
        },
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
        RegionSnapshot(
            "region_1",
            "sheet_1",
            CellRange(1, 1, 1, 1),
            "r1",
            [],
            {"logic_area_type": "unknown", "confidence": 0.35},
            False,
        )
    ]
    grouper = LLMSheetGrouper(llm_generate=failing_generate)

    grouping = grouper.group_sheet(
        sheet_info={
            "sheet_id": "sheet_1",
            "sheet_name": "Unknown",
            "sheet_index": 0,
            "max_row": 1,
            "max_col": 1,
            "merged_cells": [],
        },
        region_snapshots=snapshots,
        visual_summaries=[],
    )

    assert grouping.logic_page_name == "bill_summary_page"
    assert grouping.groups[0].region_ids == ["region_1"]
    assert "fallback" in grouping.groups[0].reason
    assert grouper.last_fallback_reason == "boom"


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
    area = LogicBuilder.build_area(
        "excel_1",
        sheet_info,
        region,
        {"logic_area_type": "fields", "confidence": 0.82},
    )

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
