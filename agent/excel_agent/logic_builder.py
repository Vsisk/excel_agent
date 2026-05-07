from __future__ import annotations

from .models import ClassificationDict, ExcelRegion, JsonDict, SheetInfoDict, gen_id


class LogicBuilder:
    """Build the logic_page and logic_area JSON structures."""

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
    def build_pages(excel_instance_id: str, sheet_list: list[SheetInfoDict]) -> list[JsonDict]:
        return [LogicBuilder.build_page(excel_instance_id, sheet_info) for sheet_info in sheet_list]

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
        }
