import re
from pathlib import Path

import pytest
from openpyxl import Workbook

from agent.excel_agent.excel_parse_handler import handle_excel_parse
from agent.excel_agent.excel_reader import ExcelReader
from agent.excel_agent.logic_builder import LogicBuilder
from agent.excel_agent.models import CellRange, ExcelRegion, gen_id
from agent.excel_agent.region_classifier import RegionClassifier
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
