from __future__ import annotations

from .models import ClassificationDict, ExcelRegion, JsonDict, RegionGroup, SheetInfoDict, gen_id


class LogicBuilder:
    """Build the logic_page and logic_area JSON structures."""

    @staticmethod
    def build_page(
        excel_instance_id: str,
        sheet_info: SheetInfoDict,
        logic_page_name: str = "bill_summary_page",
    ) -> JsonDict:
        return {
            "logic_page_id": gen_id(),
            "logic_page_name": logic_page_name,
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

    @staticmethod
    def build_grouped_area(
        excel_instance_id: str,
        sheet_info: SheetInfoDict,
        *,
        group: RegionGroup,
        region_by_id: dict[str, ExcelRegion],
        classification_by_id: dict[str, ClassificationDict] | None = None,
    ) -> JsonDict:
        first_region = region_by_id[group.region_ids[0]]
        first_range = first_region.cell_range.to_dict()
        area_type = (
            LogicBuilder._dominant_area_type(group.region_ids, classification_by_id)
            if classification_by_id
            else "unknown"
        )
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
        present = {
            classification_by_id[region_id]["logic_area_type"]
            for region_id in region_ids
            if region_id in classification_by_id
        }
        for area_type in priority:
            if area_type in present:
                return area_type
        return "unknown"
