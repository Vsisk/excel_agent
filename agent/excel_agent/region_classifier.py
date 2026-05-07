from __future__ import annotations

from .models import ClassificationDict, ExcelRegion


class RegionClassifier:
    """Legacy adapter; semantic classification is handled by the LLM pipeline."""

    @classmethod
    def classify(cls, region: ExcelRegion) -> ClassificationDict:
        return {"logic_area_type": "unknown", "confidence": 1.0}
