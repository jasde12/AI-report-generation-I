from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from statistics import pvariance
from uuid import uuid4

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt

from app.models.schemas import ExtractedSource, ExtractedTable, NormalizedDocument


@dataclass(frozen=True)
class ChartSlice:
    label: str
    value: float


class AnalysisChartGenerator:
    CHART_FONT_SIZE = 12
    FONT_CANDIDATES = [
        "Microsoft JhengHei",
        "Microsoft YaHei",
        "PingFang TC",
        "Noto Sans CJK TC",
        "SimHei",
        "Arial Unicode MS",
        "sans-serif",
    ]
    PIE_COLORS = ["#5B9BD5", "#ED7D31", "#A5A5A5", "#FFC000"]
    CATEGORY_LABELS = {
        "供電": ("電", "供電", "發電", "油量", "ups"),
        "供水": ("水", "供水", "蓄水", "水箱"),
        "供氣": ("o2", "氧", "供氧", "供氣"),
        "IT備援": ("it", "資訊", "網路", "機房"),
    }

    def __init__(self, output_dir: Path) -> None:
        self.output_dir = output_dir
        self.chart_dir = output_dir / "charts"
        self.chart_dir.mkdir(parents=True, exist_ok=True)

    def generate_pie_chart(self, normalized_document: NormalizedDocument) -> Path | None:
        chart_slices = self._extract_chart_slices(normalized_document)
        if len(chart_slices) < 2:
            return None

        output_path = self.chart_dir / f"analysis_pie_{uuid4().hex[:8]}.png"
        self._render_pie_chart(output_path, chart_slices)
        return output_path

    def _extract_chart_slices(
        self,
        normalized_document: NormalizedDocument,
    ) -> list[ChartSlice]:
        best_slices: list[ChartSlice] = []
        best_score = float("-inf")

        for source in normalized_document.sources:
            if source.type != "table":
                continue
            for table in source.tables:
                slices = self._extract_from_table(table)
                if len(slices) < 2:
                    continue
                score = self._score_slices(slices)
                if score > best_score:
                    best_score = score
                    best_slices = slices

        return best_slices

    def _extract_from_table(self, table: ExtractedTable) -> list[ChartSlice]:
        if not table.rows or not isinstance(table.rows[0], dict):
            return []

        records = [row for row in table.rows if isinstance(row, dict)]
        if not records:
            return []

        category_key = self._find_category_key(records[0])
        if category_key is None:
            return []

        candidate_keys = self._find_numeric_candidate_keys(records)
        if not candidate_keys:
            return []

        best_key = self._pick_best_value_key(candidate_keys, records)
        if best_key is None:
            return []

        aggregated: dict[str, float] = {}
        for row in records:
            category_label = self._map_category_label(row.get(category_key))
            if category_label is None:
                continue
            numeric_value = self._to_float(row.get(best_key))
            if numeric_value is None or numeric_value <= 0:
                continue
            aggregated[category_label] = numeric_value

        ordered_labels = ["供電", "供水", "供氣", "IT備援"]
        slices = [
            ChartSlice(label=label, value=aggregated[label])
            for label in ordered_labels
            if label in aggregated
        ]
        return slices

    def _find_category_key(self, sample_row: dict[str, object]) -> str | None:
        for key in sample_row:
            normalized_key = key.strip().lower()
            if any(keyword in normalized_key for keyword in ("類別", "項目", "類型", "category")):
                return key
        return None

    def _find_numeric_candidate_keys(self, records: list[dict[str, object]]) -> list[str]:
        candidates: list[str] = []
        for key in records[0]:
            if any(self._to_float(row.get(key)) is not None for row in records):
                candidates.append(key)
        return candidates

    def _pick_best_value_key(
        self,
        candidate_keys: list[str],
        records: list[dict[str, object]],
    ) -> str | None:
        best_key: str | None = None
        best_score = float("-inf")

        for key in candidate_keys:
            numeric_values = [
                value
                for row in records
                if (value := self._to_float(row.get(key))) is not None and value > 0
            ]
            if len(numeric_values) < 2:
                continue

            distinct_count = len({round(value, 4) for value in numeric_values})
            variance = pvariance(numeric_values) if len(numeric_values) > 1 else 0.0
            key_score = self._score_value_key(key, distinct_count, variance)
            if key_score > best_score:
                best_score = key_score
                best_key = key

        return best_key

    def _score_value_key(self, key: str, distinct_count: int, variance: float) -> float:
        normalized_key = key.strip().lower()
        relevance = 0.0

        if "補助後" in key:
            relevance += 1.4
        if "補助前" in key:
            relevance += 1.2
        if "天數" in key or "時數" in key or "小時" in key:
            relevance += 1.1
        if "增減" in key:
            relevance += 0.8
        if "評分" in key:
            relevance += 0.3
        if "備註" in key:
            relevance -= 1.5

        if "補助後" in key and distinct_count <= 1:
            relevance -= 0.6

        return relevance + distinct_count * 0.35 + variance * 0.1 + len(normalized_key) * 0.01

    def _score_slices(self, chart_slices: list[ChartSlice]) -> float:
        values = [chart_slice.value for chart_slice in chart_slices]
        distinct_count = len({round(value, 4) for value in values})
        variance = pvariance(values) if len(values) > 1 else 0.0
        return len(values) * 2 + distinct_count * 0.5 + variance * 0.1

    def _map_category_label(self, raw_value: object) -> str | None:
        if raw_value is None:
            return None

        normalized = str(raw_value).strip().lower()
        if not normalized:
            return None

        for label, aliases in self.CATEGORY_LABELS.items():
            if any(alias in normalized for alias in aliases):
                return label
        return None

    def _to_float(self, value: object) -> float | None:
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)

        text = str(value).strip().replace(",", "")
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None

    def _render_pie_chart(self, output_path: Path, chart_slices: list[ChartSlice]) -> None:
        plt.rcParams["font.sans-serif"] = self.FONT_CANDIDATES
        plt.rcParams["axes.unicode_minus"] = False

        labels = [chart_slice.label for chart_slice in chart_slices]
        values = [chart_slice.value for chart_slice in chart_slices]
        colors = self.PIE_COLORS[: len(chart_slices)]

        figure, axis = plt.subplots(figsize=(8.5, 4.8), dpi=200)
        axis.pie(
            values,
            labels=labels,
            autopct="%1.0f%%",
            startangle=90,
            colors=colors,
            wedgeprops={"edgecolor": "white", "linewidth": 1.5},
            textprops={"fontsize": self.CHART_FONT_SIZE},
        )
        axis.set_title("醫療韌性", fontsize=self.CHART_FONT_SIZE)
        axis.axis("equal")
        figure.tight_layout()
        figure.savefig(output_path, bbox_inches="tight")
        plt.close(figure)
