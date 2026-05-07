from __future__ import annotations

import base64
from typing import Any, Callable

from .models import RegionSnapshot, SheetInfoDict


LLMGenerate = Callable[..., dict[str, Any]]


def _default_generate_by_llm(*args: Any, **kwargs: Any) -> dict[str, Any]:
    from agent.llm.generate_by_llm import generate_by_llm

    return generate_by_llm(*args, **kwargs)


class ExcelVisualizer:
    """Collect optional visual summaries for LLM sheet grouping context."""

    def __init__(
        self,
        llm_generate: LLMGenerate = _default_generate_by_llm,
        confidence_threshold: float = 0.65,
    ):
        self.llm_generate = llm_generate
        self.confidence_threshold = confidence_threshold

    def collect_visual_summaries(
        self,
        *,
        file_uri: str,
        sheet_info: SheetInfoDict,
        region_snapshots: list[RegionSnapshot],
    ) -> dict[str, list[dict[str, Any]]]:
        summaries: list[dict[str, Any]] = []
        skipped: list[dict[str, Any]] = []

        for snapshot in region_snapshots:
            confidence = float(snapshot.rule_classification.get("confidence", 0.0))
            if confidence >= self.confidence_threshold:
                continue
            try:
                image_base64 = self._render_region_preview(snapshot)
                response = self.llm_generate(
                    "excel_visual_summary",
                    llm_name="vl",
                    lang="zh",
                    target_id=snapshot.region_id,
                    image_base64=image_base64,
                    image_mime_type="image/png",
                )
                summaries.append(
                    {
                        "target_id": str(response.get("target_id", snapshot.region_id)),
                        "target_type": "cell_region",
                        "sheet_id": sheet_info["sheet_id"],
                        "summary": str(response.get("summary", "")),
                        "confidence": float(response.get("confidence", 0.0)),
                    }
                )
            except Exception as exc:
                skipped.append(
                    {
                        "target_id": snapshot.region_id,
                        "target_type": "cell_region",
                        "sheet_id": sheet_info["sheet_id"],
                        "reason": str(exc),
                    }
                )

        skipped.extend(self._embedded_image_skips(file_uri, sheet_info))
        return {"summaries": summaries, "skipped": skipped}

    @staticmethod
    def _render_region_preview(snapshot: RegionSnapshot) -> str:
        # Minimal deterministic PNG-like payload for testable VL plumbing.
        # A richer renderer can replace this without changing call sites.
        pseudo_png = b"\x89PNG\r\n\x1a\n" + snapshot.markdown.encode("utf-8")
        return base64.b64encode(pseudo_png).decode("ascii")

    @staticmethod
    def _embedded_image_skips(file_uri: str, sheet_info: SheetInfoDict) -> list[dict[str, Any]]:
        return []
