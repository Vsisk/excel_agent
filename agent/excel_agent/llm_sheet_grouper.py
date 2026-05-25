from __future__ import annotations

import json
from typing import Any, Callable

from .models import (
    ALLOWED_LOGIC_PAGE_NAMES,
    RegionGroup,
    RegionSnapshot,
    SheetGrouping,
    SheetInfoDict,
)


LLMGenerate = Callable[..., dict[str, Any]]


def _default_generate_by_llm(*args: Any, **kwargs: Any) -> dict[str, Any]:
    from agent.llm.generate_by_llm import generate_by_llm

    return generate_by_llm(*args, **kwargs)


class LLMSheetGrouper:
    """Call the base LLM once per sheet and normalize minimal grouping output."""

    def __init__(self, llm_generate: LLMGenerate = _default_generate_by_llm):
        self.llm_generate = llm_generate
        self.last_fallback_reason: str | None = None
        self.last_llm_used = False

    def group_sheet(
        self,
        *,
        sheet_info: SheetInfoDict,
        region_snapshots: list[RegionSnapshot],
        grouping_memory_matches: list[dict[str, Any]] | None = None,
    ) -> SheetGrouping:
        self.last_fallback_reason = None
        self.last_llm_used = False
        try:
            payload = self._payload(
                sheet_info,
                region_snapshots,
                grouping_memory_matches or [],
            )
            response = self.llm_generate(
                "excel_sheet_grouping",
                llm_name="base",
                lang="zh",
                sheet_payload=json.dumps(payload, ensure_ascii=False),
            )
            grouping = self._normalize(response, region_snapshots)
            self.last_llm_used = True
            return grouping
        except Exception as exc:
            self.last_fallback_reason = str(exc)
            return self.fallback(region_snapshots, reason=f"fallback: {exc}")

    @staticmethod
    def fallback(region_snapshots: list[RegionSnapshot], reason: str) -> SheetGrouping:
        return SheetGrouping(
            logic_page_name="bill_summary_page",
            groups=[
                RegionGroup(region_ids=[snapshot.region_id], reason=reason)
                for snapshot in region_snapshots
            ],
        )

    @staticmethod
    def _payload(
        sheet_info: SheetInfoDict,
        region_snapshots: list[RegionSnapshot],
        grouping_memory_matches: list[dict[str, Any]],
    ) -> dict[str, Any]:
        return {
            "sheet": {
                "sheet_id": sheet_info["sheet_id"],
                "sheet_name": sheet_info["sheet_name"],
                "sheet_index": sheet_info["sheet_index"],
            },
            "grouping_task": (
                "Merge rule-split regions that are fragments of the same logical table or section. "
                "Regions can be split by layout gaps, merged cells, page formatting, or visual spacing. "
                "Return groups of region_id values in reading order; keep unrelated tables separate."
            ),
            "allowed_logic_page_names": list(ALLOWED_LOGIC_PAGE_NAMES),
            "regions": [
                {
                    "region_id": snapshot.region_id,
                    "bbox": {
                        "left": snapshot.cell_range.start_col,
                        "right": snapshot.cell_range.end_col,
                        "top": snapshot.cell_range.start_row,
                        "bottom": snapshot.cell_range.end_row,
                    },
                    "table_md": snapshot.markdown,
                    "truncated": snapshot.truncated,
                }
                for snapshot in region_snapshots
            ],
            "grouping_memory_matches": grouping_memory_matches,
        }

    @staticmethod
    def _normalize(response: dict[str, Any], region_snapshots: list[RegionSnapshot]) -> SheetGrouping:
        region_order = [snapshot.region_id for snapshot in region_snapshots]
        valid_ids = set(region_order)
        page_name = response.get("logic_page_name")
        if page_name not in ALLOWED_LOGIC_PAGE_NAMES:
            page_name = "bill_summary_page"

        seen: set[str] = set()
        groups: list[RegionGroup] = []
        for raw_group in response.get("groups", []):
            if not isinstance(raw_group, dict):
                continue
            ids = raw_group.get("region_ids", [])
            if not isinstance(ids, list):
                continue
            deduped = [
                region_id
                for region_id in region_order
                if region_id in ids and region_id not in seen and region_id in valid_ids
            ]
            if not deduped:
                continue
            seen.update(deduped)
            reason = raw_group.get("reason")
            groups.append(RegionGroup(region_ids=deduped, reason=str(reason or "llm grouping")))

        for region_id in region_order:
            if region_id not in seen:
                groups.append(RegionGroup(region_ids=[region_id], reason="fallback: omitted by llm"))

        return SheetGrouping(logic_page_name=page_name, groups=groups)
