from __future__ import annotations

import re

from .models import ClassificationDict, ExcelRegion


class RegionClassifier:
    """Rule-based classifier with the same surface a future LLM classifier can keep."""

    HEADER_KEYWORDS = {
        "amount",
        "fee",
        "price",
        "qty",
        "quantity",
        "item",
        "description",
        "subtotal",
        "total",
        "金额",
        "费用",
        "数量",
        "项目",
        "名称",
    }

    @classmethod
    def classify(cls, region: ExcelRegion) -> ClassificationDict:
        if region.row_count <= 5 and cls._looks_like_fields(region):
            return {"logic_area_type": "fields", "confidence": 0.82}

        if cls._has_clear_header(region):
            area_type = "fee_table" if cls._looks_like_fee_table(region) else "detail_table"
            return {"logic_area_type": area_type, "confidence": 0.78}

        if cls._looks_like_plain_text(region):
            return {"logic_area_type": "plain_text", "confidence": 0.74}

        return {"logic_area_type": "unknown", "confidence": 0.35}

    @classmethod
    def _looks_like_fields(cls, region: ExcelRegion) -> bool:
        if region.col_count != 2 or not region.raw_text:
            return False
        short_lines = [line for line in region.raw_text if 2 <= len(line.split()) <= 8]
        return len(short_lines) >= max(1, len(region.raw_text) // 2)

    @classmethod
    def _has_clear_header(cls, region: ExcelRegion) -> bool:
        if region.row_count < 2 or region.col_count < 2 or not region.raw_text:
            return False
        first_line = region.raw_text[0].lower()
        matches = sum(1 for keyword in cls.HEADER_KEYWORDS if keyword in first_line)
        return matches >= 2 or (matches >= 1 and region.col_count >= 3)

    @classmethod
    def _looks_like_fee_table(cls, region: ExcelRegion) -> bool:
        joined = " ".join(region.raw_text).lower()
        return any(token in joined for token in ["amount", "fee", "price", "total", "金额", "费用"])

    @classmethod
    def _looks_like_plain_text(cls, region: ExcelRegion) -> bool:
        joined = " ".join(region.raw_text)
        words = re.findall(r"\S+", joined)
        return region.col_count <= 2 and len(words) >= 10
