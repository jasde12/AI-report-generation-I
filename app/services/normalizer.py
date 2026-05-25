from __future__ import annotations

import re
from typing import Any

from app.models.schemas import ExtractedSource, ExtractedTable, NormalizedDocument, NormalizedFact


class Normalizer:
    def normalize(
        self,
        project_name: str,
        hospital_name: str | None,
        template_id: str,
        sources: list[ExtractedSource],
    ) -> NormalizedDocument:
        warnings: list[str] = []
        facts: list[NormalizedFact] = []
        resolved_hospital_name = (hospital_name or "").strip() or self._infer_hospital_name(sources)

        for source in sources:
            warnings.extend(f"{source.source_name}: {warning}" for warning in source.warnings)
            for table in source.tables:
                warnings.extend(
                    f"{source.source_name}::{table.title or 'Untitled Table'}: {warning}"
                    for warning in table.warnings
                )
            facts.extend(self._source_to_facts(source))

        return NormalizedDocument(
            project_name=project_name,
            hospital_name=resolved_hospital_name or None,
            template_id=template_id,
            sources=sources,
            facts=facts,
            warnings=warnings,
        )

    def build_preview(
        self,
        normalized_document: NormalizedDocument,
        data_url_preview_chars: int = 120,
    ) -> dict[str, Any]:
        preview = normalized_document.model_dump(mode="json")
        for source in preview.get("sources", []):
            data_url = source.get("data_url")
            if data_url:
                source["data_url"] = self._truncate_data_url(data_url, data_url_preview_chars)
        return preview

    def _source_to_facts(self, source: ExtractedSource) -> list[NormalizedFact]:
        facts: list[NormalizedFact] = []

        if source.blocks:
            facts.extend(self._blocks_to_facts(source))
        else:
            facts.extend(self._legacy_text_facts(source))

        if source.text:
            for block in self._split_text_blocks(source.text):
                facts.append(
                    NormalizedFact(
                        source_name=source.source_name,
                        fact_type="text_block",
                        content=block,
                    )
                )

        if source.image_analysis:
            facts.append(
                NormalizedFact(
                    source_name=source.source_name,
                    fact_type="image_summary",
                    content=source.image_analysis.summary,
                )
            )
            for item in source.image_analysis.detected_items:
                facts.append(
                    NormalizedFact(
                        source_name=source.source_name,
                        fact_type="image_item",
                        content=item,
                    )
                )
            for checkbox in source.image_analysis.checkbox_values:
                details = f"{checkbox.field_name}: {checkbox.value}"
                if checkbox.confidence_note:
                    details = f"{details} ({checkbox.confidence_note})"
                facts.append(
                    NormalizedFact(
                        source_name=source.source_name,
                        fact_type="checkbox_value",
                        content=details,
                    )
                )

        return facts

    def _blocks_to_facts(self, source: ExtractedSource) -> list[NormalizedFact]:
        facts: list[NormalizedFact] = []
        tables_by_index = {
            index: table
            for index, table in enumerate(source.tables, start=1)
        }

        for block in source.blocks:
            if block.block_type == "paragraph" and block.text:
                facts.append(
                    NormalizedFact(
                        source_name=source.source_name,
                        fact_type="paragraph",
                        content=block.text,
                        block_order=block.order,
                    )
                )
                continue

            if block.block_type != "table":
                continue

            table = tables_by_index.get(block.table_index or 0)
            table_title = block.table_title or (table.title if table else "Untitled Table")
            if block.table_markdown:
                facts.append(
                    NormalizedFact(
                        source_name=source.source_name,
                        fact_type="table_markdown",
                        content=f"{table_title}\n{block.table_markdown}",
                        block_order=block.order,
                    )
                )

            if table is None:
                continue

            table_structure = self._build_table_structure(table)
            if table_structure:
                facts.append(
                    NormalizedFact(
                        source_name=source.source_name,
                        fact_type="table_structure",
                        content=f"{table_title}\n{table_structure}",
                        block_order=block.order,
                    )
                )

            table_semantic = self._build_table_semantic(table)
            if table_semantic:
                facts.append(
                    NormalizedFact(
                        source_name=source.source_name,
                        fact_type="table_semantic",
                        content=f"{table_title}\n{table_semantic}",
                        block_order=block.order,
                    )
                )

            table_questions = self._build_table_questions(table)
            if table_questions:
                facts.append(
                    NormalizedFact(
                        source_name=source.source_name,
                        fact_type="table_questions",
                        content=f"{table_title}\n{table_questions}",
                        block_order=block.order,
                    )
                )

        return facts

    def _legacy_text_facts(self, source: ExtractedSource) -> list[NormalizedFact]:
        facts: list[NormalizedFact] = []

        for paragraph in source.paragraphs:
            facts.append(
                NormalizedFact(
                    source_name=source.source_name,
                    fact_type="paragraph",
                    content=paragraph,
                )
            )

        for table in source.tables:
            table_title = table.title or "Untitled Table"
            facts.append(
                NormalizedFact(
                    source_name=source.source_name,
                    fact_type="table_markdown",
                    content=f"{table_title}\n{table.markdown}",
                )
            )

            table_structure = self._build_table_structure(table)
            if table_structure:
                facts.append(
                    NormalizedFact(
                        source_name=source.source_name,
                        fact_type="table_structure",
                        content=f"{table_title}\n{table_structure}",
                    )
                )

            table_semantic = self._build_table_semantic(table)
            if table_semantic:
                facts.append(
                    NormalizedFact(
                        source_name=source.source_name,
                        fact_type="table_semantic",
                        content=f"{table_title}\n{table_semantic}",
                    )
                )

            table_questions = self._build_table_questions(table)
            if table_questions:
                facts.append(
                    NormalizedFact(
                        source_name=source.source_name,
                        fact_type="table_questions",
                        content=f"{table_title}\n{table_questions}",
                    )
                )

        return facts

    def _build_table_structure(self, table: ExtractedTable) -> str:
        details: list[str] = []

        if table.row_count or table.column_count:
            details.append(f"row_count={table.row_count}, column_count={table.column_count}")

        if table.table_kind:
            details.append(f"table_kind={table.table_kind}")

        if table.header_depth:
            details.append(f"header_depth={table.header_depth}")

        if table.header_rows:
            header_rows_text = " | ".join(
                " / ".join(cell for cell in row if cell)
                for row in table.header_rows[:3]
            )
            if header_rows_text:
                details.append(f"header_rows={header_rows_text}")

        if table.warnings:
            details.append(f"warnings={' ; '.join(table.warnings)}")

        if table.merged_ranges:
            merged_ranges_text = ", ".join(
                (
                    f"R{merged_range.start_row}C{merged_range.start_column}"
                    f"-R{merged_range.end_row}C{merged_range.end_column}"
                )
                for merged_range in table.merged_ranges[:20]
            )
            if len(table.merged_ranges) > 20:
                merged_ranges_text = (
                    f"{merged_ranges_text}, ... total={len(table.merged_ranges)}"
                )
            details.append(f"merged_ranges={merged_ranges_text}")

        if table.cells:
            cell_descriptions: list[str] = []
            for cell in table.cells[:40]:
                span_text = ""
                if cell.row_span > 1 or cell.column_span > 1:
                    span_text = f"[row_span={cell.row_span}, column_span={cell.column_span}]"
                cell_text = cell.text.replace("\n", " / ").strip() or "(blank)"
                cell_descriptions.append(
                    f"R{cell.row_index}C{cell.column_index}{span_text}={cell_text}"
                )

            cell_summary = "; ".join(cell_descriptions)
            if len(table.cells) > 40:
                cell_summary = f"{cell_summary}; ... total={len(table.cells)}"
            details.append(f"cells={cell_summary}")

        return "\n".join(details)

    def _build_table_semantic(self, table: ExtractedTable) -> str:
        if not table.semantic_lines:
            return ""

        priority_markers = ("selected_values", "checkbox_options")
        priority_lines = [
            line
            for line in table.semantic_lines
            if any(marker in line for marker in priority_markers)
        ]
        remaining_lines = [
            line for line in table.semantic_lines if line not in priority_lines
        ]
        lines = (priority_lines + remaining_lines)[:80]
        if len(table.semantic_lines) > 80:
            lines.append(f"... total={len(table.semantic_lines)}")
        return "\n".join(lines)

    def _build_table_questions(self, table: ExtractedTable) -> str:
        if not table.questions:
            return ""

        lines: list[str] = []
        for question in table.questions[:80]:
            options = "; ".join(
                f"{option.label}: {'selected' if option.selected else 'not_selected'}"
                for option in question.options[:30]
            )
            selected_values = "; ".join(question.selected_values)
            details = (
                f"{question.question_id} | {question.question} | "
                f"selection_type={question.selection_type} | options={options}"
            )
            if selected_values:
                details = f"{details} | selected_values={selected_values}"
            lines.append(details)

        if len(table.questions) > 80:
            lines.append(f"... total={len(table.questions)}")
        return "\n".join(lines)

    def _split_text_blocks(self, text: str) -> list[str]:
        blocks: list[str] = []
        for line in text.splitlines():
            cleaned = line.strip()
            if cleaned:
                blocks.append(cleaned)
        return blocks

    def _truncate_data_url(self, data_url: str, preview_chars: int) -> str:
        if len(data_url) <= preview_chars:
            return data_url
        return f"{data_url[:preview_chars]}...(truncated)"

    def _infer_hospital_name(self, sources: list[ExtractedSource]) -> str:
        explicit_name_pattern = re.compile(r"醫院名稱[:：]\s*([^\n，。；（）()]{1,30}醫院)")
        ministry_name_pattern = re.compile(r"(衛生福利部[^\s，。；（）()]{1,24}醫院)")
        generic_name_pattern = re.compile(r"([^\s，。；（）()]{1,24}醫院)")

        explicit_candidates: list[str] = []
        ministry_candidates: list[str] = []
        generic_candidates: list[str] = []

        def looks_like_hospital_name(value: str) -> bool:
            if not value or "醫院" not in value:
                return False
            if any(keyword in value for keyword in ("年度", "計畫", "報告", "附件", "責任醫院", "急救")):
                return value.endswith("醫院") and "衛生福利部" in value
            return True

        def collect(text: str) -> None:
            for match in explicit_name_pattern.findall(text):
                cleaned = match.strip()
                if cleaned and looks_like_hospital_name(cleaned):
                    explicit_candidates.append(cleaned)
            for match in ministry_name_pattern.findall(text):
                cleaned = match.strip()
                if cleaned and looks_like_hospital_name(cleaned):
                    ministry_candidates.append(cleaned)
            for match in generic_name_pattern.findall(text):
                cleaned = match.strip()
                if cleaned and looks_like_hospital_name(cleaned):
                    generic_candidates.append(cleaned)

        for source in sources:
            for paragraph in source.paragraphs:
                collect(paragraph)
            if source.text:
                collect(source.text)
            for table in source.tables:
                for cell in table.cells:
                    collect(cell.text)

        for candidates in (explicit_candidates, ministry_candidates, generic_candidates):
            if candidates:
                return candidates[0]

        return ""
