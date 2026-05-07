from __future__ import annotations

from typing import Any

from .models import ClassificationDict, ExcelRegion, RegionSnapshot


class RegionMarkdownBuilder:
    """Convert a rule-split region into compact Markdown for LLM grouping."""

    @staticmethod
    def build_region_snapshot(
        *,
        region_id: str,
        region: ExcelRegion,
        rows: list[list[Any]],
        classification: ClassificationDict,
        max_instance_rows: int = 10,
    ) -> RegionSnapshot:
        normalized = [[RegionMarkdownBuilder._stringify(cell) for cell in row] for row in rows]
        normalized = [row for row in normalized if any(cell for cell in row)]
        truncated = False

        if RegionMarkdownBuilder._is_table(normalized):
            header = normalized[0]
            body = normalized[1:]
            if len(body) > max_instance_rows:
                body = body[:max_instance_rows]
                truncated = True
            markdown = RegionMarkdownBuilder._table(header, body)
        else:
            lines = ["- " + " ".join(cell for cell in row if cell) for row in normalized]
            markdown = "\n".join(lines)

        if truncated:
            markdown = f"{markdown}\n\n... truncated after {max_instance_rows} rows"

        return RegionSnapshot(
            region_id=region_id,
            sheet_id=region.sheet_id,
            cell_range=region.cell_range,
            markdown=markdown,
            raw_text=region.raw_text,
            rule_classification=classification,
            truncated=truncated,
        )

    @staticmethod
    def _is_table(rows: list[list[str]]) -> bool:
        return len(rows) > 1 and max((len(row) for row in rows), default=0) > 1

    @staticmethod
    def _table(header: list[str], body: list[list[str]]) -> str:
        width = max(len(header), *(len(row) for row in body)) if body else len(header)
        padded_header = RegionMarkdownBuilder._pad(header, width)
        lines = [
            "| " + " | ".join(padded_header) + " |",
            "| " + " | ".join("---" for _ in range(width)) + " |",
        ]
        for row in body:
            lines.append("| " + " | ".join(RegionMarkdownBuilder._pad(row, width)) + " |")
        return "\n".join(lines)

    @staticmethod
    def _pad(row: list[str], width: int) -> list[str]:
        return row + ["" for _ in range(width - len(row))]

    @staticmethod
    def _stringify(value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()
