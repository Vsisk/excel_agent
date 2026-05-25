from __future__ import annotations

from dataclasses import dataclass, field
from functools import lru_cache
import math
import os
from typing import Callable

from .models import RegionGroup, RegionSnapshot, SheetGrouping, SheetInfoDict


EmbeddingGenerate = Callable[[str], list[float]]

DEFAULT_EMBEDDING_MODEL = "BAAI/bge-m3"
EMBEDDING_MODEL_ENV = "EXCEL_AGENT_EMBEDDING_MODEL"
EMBEDDING_STRICT_ENV = "EXCEL_AGENT_EMBEDDING_STRICT"


def default_embedding_generate(text: str) -> list[float]:
    model_name = os.getenv(EMBEDDING_MODEL_ENV, DEFAULT_EMBEDDING_MODEL)
    try:
        import torch
        import torch.nn.functional as F
    except Exception as exc:
        raise RuntimeError(f"embedding unavailable: {exc}") from exc

    tokenizer, model = _load_embedding_model(model_name)
    encoded = tokenizer(
        text,
        padding=True,
        truncation=True,
        max_length=8192,
        return_tensors="pt",
    )
    with torch.no_grad():
        output = model(**encoded)
        embeddings = _mean_pool(output.last_hidden_state, encoded["attention_mask"])
        embeddings = F.normalize(embeddings, p=2, dim=1)
    return embeddings[0].tolist()


@lru_cache(maxsize=2)
def _load_embedding_model(model_name: str):
    try:
        from transformers import AutoModel, AutoTokenizer
    except Exception as exc:
        raise RuntimeError(f"embedding unavailable: {exc}") from exc

    tokenizer = AutoTokenizer.from_pretrained(model_name)
    model = AutoModel.from_pretrained(model_name)
    model.eval()
    return tokenizer, model


def _mean_pool(token_embeddings, attention_mask):
    import torch

    mask = attention_mask.unsqueeze(-1).expand(token_embeddings.size()).float()
    return torch.sum(token_embeddings * mask, 1) / torch.clamp(mask.sum(1), min=1e-9)


@dataclass
class GroupTemplate:
    template_id: str
    template_text: str
    embedding: list[float]
    support_count: int = 1
    sample_sheet_names: list[str] = field(default_factory=list)
    group_reason: str = ""

    def preview(self) -> str:
        first_line = self.template_text.splitlines()[0] if self.template_text else self.group_reason
        return first_line[:180]


class WorkbookGroupingMemory:
    """Per-workbook group template memory backed by embedding similarity."""

    def __init__(
        self,
        *,
        embedding_generate: EmbeddingGenerate | None = None,
        top_k: int = 5,
        max_templates: int = 5,
        max_samples_per_template: int = 2,
        merge_threshold: float = 0.9,
        strong_match_threshold: float = 0.75,
        strict: bool | None = None,
    ):
        self.embedding_generate = embedding_generate or default_embedding_generate
        self.top_k = top_k
        self.max_templates = max_templates
        self.max_samples_per_template = max_samples_per_template
        self.merge_threshold = merge_threshold
        self.strong_match_threshold = strong_match_threshold
        self.strict = strict if strict is not None else os.getenv(EMBEDDING_STRICT_ENV) == "1"
        self.enabled = True
        self.used = False
        self.fallback_reason: str | None = None
        self.matches_by_sheet: list[dict] = []
        self.warnings: list[dict] = []
        self._templates: list[GroupTemplate] = []
        self._next_id = 1

    @property
    def template_count(self) -> int:
        return len(self._templates)

    def retrieve_for_sheet(
        self,
        *,
        sheet_info: SheetInfoDict,
        region_snapshots: list[RegionSnapshot],
        visual_summaries: list[dict],
    ) -> list[dict]:
        if not self.enabled or not self._templates:
            return []

        query_text = self._sheet_query_text(sheet_info, region_snapshots, visual_summaries)
        query_embedding = self._embed(query_text)
        if query_embedding is None:
            return []

        scored = [
            (self._cosine_similarity(query_embedding, template.embedding), template)
            for template in self._templates
        ]
        scored.sort(key=lambda item: item[0], reverse=True)
        matches = [
            self._match_summary(template, similarity)
            for similarity, template in scored[: self.top_k]
            if similarity > 0
        ]
        if matches:
            self.used = True
        return matches

    def remember_sheet_grouping(
        self,
        *,
        sheet_info: SheetInfoDict,
        region_snapshots: list[RegionSnapshot],
        visual_summaries: list[dict],
        grouping: SheetGrouping,
    ) -> None:
        if not self.enabled:
            return

        snapshot_by_id = {snapshot.region_id: snapshot for snapshot in region_snapshots}
        for group in grouping.groups:
            text = self._group_template_text(sheet_info, group, snapshot_by_id, visual_summaries)
            embedding = self._embed(text)
            if embedding is None:
                return
            match = self._best_template(embedding)
            if match and match[0] >= self.merge_threshold:
                template = match[1]
                template.support_count += 1
                if sheet_info["sheet_name"] not in template.sample_sheet_names:
                    template.sample_sheet_names.append(sheet_info["sheet_name"])
                    del template.sample_sheet_names[self.max_samples_per_template :]
                continue

            self._templates.append(
                GroupTemplate(
                    template_id=f"template_{self._next_id}",
                    template_text=text,
                    embedding=embedding,
                    sample_sheet_names=[sheet_info["sheet_name"]],
                    group_reason=group.reason,
                )
            )
            self._next_id += 1
            self._templates = self._templates[-self.max_templates :]

    def record_sheet_matches(self, sheet_info: SheetInfoDict, matches: list[dict]) -> None:
        self.matches_by_sheet.append(
            {
                "sheet_id": sheet_info["sheet_id"],
                "sheet_name": sheet_info["sheet_name"],
                "matches": matches,
            }
        )

    def record_consistency_warnings(
        self,
        sheet_info: SheetInfoDict,
        matches: list[dict],
        grouping: SheetGrouping,
    ) -> None:
        if not matches or not grouping.groups:
            return
        if len(grouping.groups) == 1:
            return
        for match in matches:
            if match["similarity"] >= self.strong_match_threshold:
                self.warnings.append(
                    {
                        "sheet_id": sheet_info["sheet_id"],
                        "sheet_name": sheet_info["sheet_name"],
                        "matched_template_id": match["template_id"],
                        "similarity": match["similarity"],
                        "message": "High-similarity grouping memory was available; review current multi-group output for consistency.",
                    }
                )
                return

    def _embed(self, text: str) -> list[float] | None:
        try:
            return self.embedding_generate(text)
        except Exception as exc:
            self.fallback_reason = str(exc)
            self.enabled = False
            if self.strict:
                raise
            return None

    def _best_template(self, embedding: list[float]) -> tuple[float, GroupTemplate] | None:
        if not self._templates:
            return None
        scored = [(self._cosine_similarity(embedding, template.embedding), template) for template in self._templates]
        return max(scored, key=lambda item: item[0])

    @staticmethod
    def _cosine_similarity(left: list[float], right: list[float]) -> float:
        if not left or not right or len(left) != len(right):
            return 0.0
        dot = sum(a * b for a, b in zip(left, right))
        left_norm = math.sqrt(sum(a * a for a in left))
        right_norm = math.sqrt(sum(b * b for b in right))
        if left_norm == 0 or right_norm == 0:
            return 0.0
        return round(dot / (left_norm * right_norm), 6)

    def _match_summary(self, template: GroupTemplate, similarity: float) -> dict:
        return {
            "template_id": template.template_id,
            "similarity": round(similarity, 6),
            "support_count": template.support_count,
            "sample_sheet_names": list(template.sample_sheet_names),
            "template_preview": template.preview(),
        }

    @staticmethod
    def _sheet_query_text(
        sheet_info: SheetInfoDict,
        region_snapshots: list[RegionSnapshot],
        visual_summaries: list[dict],
    ) -> str:
        parts = [f"sheet={sheet_info['sheet_name']}"]
        parts.extend(WorkbookGroupingMemory._region_text(snapshot) for snapshot in region_snapshots)
        parts.extend(WorkbookGroupingMemory._visual_text(visual) for visual in visual_summaries)
        return "\n".join(part for part in parts if part)

    @staticmethod
    def _group_template_text(
        sheet_info: SheetInfoDict,
        group: RegionGroup,
        snapshot_by_id: dict[str, RegionSnapshot],
        visual_summaries: list[dict],
    ) -> str:
        snapshots = [snapshot_by_id[region_id] for region_id in group.region_ids if region_id in snapshot_by_id]
        visual_by_target = {visual.get("target_id"): visual for visual in visual_summaries}
        parts = [
            f"{group.reason} | regions={len(snapshots)} | "
            + " ".join(WorkbookGroupingMemory._keywords(snapshot.markdown) for snapshot in snapshots)
        ]
        parts.append(f"source_sheet={sheet_info['sheet_name']}")
        for snapshot in snapshots:
            parts.append(WorkbookGroupingMemory._region_text(snapshot))
            visual = visual_by_target.get(snapshot.region_id)
            if visual:
                parts.append(WorkbookGroupingMemory._visual_text(visual))
        return "\n".join(part for part in parts if part)

    @staticmethod
    def _region_text(snapshot: RegionSnapshot) -> str:
        cell_range = snapshot.cell_range
        row_count = cell_range.end_row - cell_range.start_row + 1
        col_count = cell_range.end_col - cell_range.start_col + 1
        size = f"rows={row_count}, cols={col_count}"
        markdown_preview = snapshot.markdown.replace("\n", " ")[:500]
        keywords = WorkbookGroupingMemory._keywords(snapshot.markdown)
        return f"{snapshot.region_id}: {size}; keywords={keywords}; {markdown_preview}"

    @staticmethod
    def _visual_text(visual: dict) -> str:
        summary = str(visual.get("summary", ""))[:300]
        return f"visual {visual.get('target_id')}: {summary}"

    @staticmethod
    def _keywords(text: str) -> str:
        words = []
        for token in text.replace("|", " ").replace("-", " ").split():
            cleaned = token.strip().lower()
            if len(cleaned) >= 2 and cleaned not in words:
                words.append(cleaned)
            if len(words) >= 12:
                break
        return " ".join(words)
