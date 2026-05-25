from __future__ import annotations

import re
from pathlib import Path
from typing import Iterator

from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl, CT_Tc
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

from app.models.schemas import (
    ExtractedQuestion,
    ExtractedSource,
    ExtractedTable,
    MergedCellRange,
    QuestionOption,
    SourceBlock,
    TableCell,
)


class DocxExtractor:
    CHECKED_SYMBOLS = frozenset({"■", "☑", "", "✓"})
    UNCHECKED_SYMBOLS = frozenset({"□", "☐", "", "○"})
    TEXT_CHECKED_SYMBOLS = frozenset({"v", "V"})
    CHECKBOX_MARKER_PATTERN = re.compile(
        r"([■□☑☐✓○]|(?<![A-Za-z0-9])[vV](?![A-Za-z0-9]))"
    )
    SYMBOL_MAP: dict[tuple[str, str], str] = {
        ("wingdings", "F06E"): "■",
        ("wingdings", "F06F"): "□",
        ("wingdings", "F0E8"): "○",
        ("wingdings", "F08C"): "○",
        ("wingdings", "F08D"): "○",
        ("wingdings", "F08E"): "○",
        ("wingdings", "F08F"): "○",
        ("wingdings 2", "F051"): "□",
        ("wingdings 2", "F052"): "■",
    }
    DEFAULT_SYMBOL_MAP: dict[str, str] = {
        "F06E": "■",
        "F06F": "□",
        "F0E8": "○",
        "F08C": "○",
        "F08D": "○",
        "F08E": "○",
        "F08F": "○",
        "F051": "□",
        "F052": "■",
    }
    TEXT_SYMBOL_MAP: dict[str, str] = {
        chr(0xF06E): "■",
        chr(0xF06F): "□",
        chr(0xF0E8): "○",
        chr(0xF08C): "○",
        chr(0xF08D): "○",
        chr(0xF08E): "○",
        chr(0xF08F): "○",
        chr(0xF051): "□",
        chr(0xF052): "■",
    }

    def extract(self, file_path: Path) -> ExtractedSource:
        try:
            document = Document(file_path)
        except Exception as exc:
            raise ValueError(f"無法讀取 DOCX 檔案: {file_path.name}") from exc

        paragraphs: list[str] = []
        tables: list[ExtractedTable] = []
        blocks: list[SourceBlock] = []
        table_index = 0

        for block_order, block in enumerate(self._iter_document_blocks(document), start=1):
            if isinstance(block, Paragraph):
                text = self._extract_paragraph_text(block)
                if not text:
                    continue
                paragraphs.append(text)
                blocks.append(
                    SourceBlock(
                        order=block_order,
                        block_type="paragraph",
                        text=text,
                    )
                )
                continue

            table_index += 1
            extracted_table = self._extract_table(block, table_index=table_index)
            if extracted_table is None:
                continue

            tables.append(extracted_table)
            blocks.append(
                SourceBlock(
                    order=block_order,
                    block_type="table",
                    table_index=table_index,
                    table_title=extracted_table.title,
                    table_markdown=extracted_table.markdown,
                )
            )

        return ExtractedSource(
            type="docx",
            source_name=file_path.name,
            paragraphs=paragraphs,
            tables=tables,
            blocks=blocks,
            metadata={"path": str(file_path)},
        )

    def _iter_document_blocks(self, document: DocxDocument) -> Iterator[Paragraph | Table]:
        for child in document.element.body.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, document._body)
            elif isinstance(child, CT_Tbl):
                yield Table(child, document._body)

    def _extract_table(self, table: Table, table_index: int) -> ExtractedTable | None:
        rows, cells, row_count, column_count = self._extract_table_grid(table)
        if not rows and not cells:
            return None

        padded_rows = [row + [""] * (column_count - len(row)) for row in rows]
        merged_ranges = [
            MergedCellRange(
                start_row=cell.row_index,
                start_column=cell.column_index,
                end_row=cell.row_index + cell.row_span - 1,
                end_column=cell.column_index + cell.column_span - 1,
            )
            for cell in cells
            if cell.row_span > 1 or cell.column_span > 1
        ]
        table_kind = self._detect_table_kind(padded_rows, cells, row_count, column_count)
        header_rows = self._infer_header_rows(table_kind, padded_rows)
        semantic_lines = self._build_semantic_lines(
            table_kind=table_kind,
            rows=padded_rows,
            cells=cells,
        )
        questions = self._build_structured_questions(
            table_kind=table_kind,
            rows=padded_rows,
        )
        warnings = self._build_table_warnings(
            table_kind=table_kind,
            merged_ranges=merged_ranges,
            semantic_lines=semantic_lines,
        )
        columns = header_rows[-1] if header_rows else (padded_rows[0] if padded_rows else [])

        return ExtractedTable(
            title=f"Table {table_index}",
            table_kind=table_kind,
            columns=columns,
            rows=padded_rows,
            raw_rows=padded_rows,
            header_rows=header_rows,
            semantic_lines=semantic_lines,
            questions=questions,
            markdown=self._rows_to_markdown(padded_rows),
            row_count=row_count,
            column_count=column_count,
            cells=cells,
            merged_ranges=merged_ranges,
            header_depth=len(header_rows),
            warnings=warnings,
        )

    def _detect_table_kind(
        self,
        rows: list[list[str]],
        cells: list[TableCell],
        row_count: int,
        column_count: int,
    ) -> str:
        if not rows:
            return "generic_matrix"

        non_empty_cells = [cell for cell in cells if cell.text.strip()]
        if len(non_empty_cells) == 1 and row_count == 1 and column_count == 1:
            return "metadata_block"

        if row_count >= 4 and column_count >= 3:
            header_row = rows[2]
            header_text = " ".join(cell for cell in header_row if cell)
            if (
                "項目" in header_text
                and any("現況" in cell for cell in header_row)
                and any("建議" in cell for cell in header_row)
            ):
                return "diagnostic_matrix"

        if self._looks_like_header_row(rows[0]):
            return "questionnaire_matrix"

        if column_count == 2 and row_count >= 2:
            compact_row_labels = sum(
                1
                for row in rows[1: min(len(rows), 5)]
                if row
                and row[0].strip()
                and len(self._to_inline_text(row[0])) <= 40
            )
            if compact_row_labels >= 1:
                return "questionnaire_matrix"

            first_left_cell = self._to_inline_text(rows[0][0]) if rows[0] else ""
            if len(first_left_cell) > 80:
                return "narrative_matrix"

        return "generic_matrix"

    def _infer_header_rows(self, table_kind: str, rows: list[list[str]]) -> list[list[str]]:
        if not rows:
            return []

        if table_kind == "diagnostic_matrix" and len(rows) >= 3:
            return [rows[2]]

        if table_kind == "questionnaire_matrix" and self._looks_like_header_row(rows[0]):
            return [rows[0]]

        return []

    def _build_semantic_lines(
        self,
        table_kind: str,
        rows: list[list[str]],
        cells: list[TableCell],
    ) -> list[str]:
        if table_kind == "metadata_block" and cells:
            return self._unique_lines(self._build_metadata_block_lines(cells[0].text))

        if table_kind == "diagnostic_matrix":
            return self._unique_lines(self._build_diagnostic_matrix_lines(rows))

        if table_kind == "questionnaire_matrix":
            return self._unique_lines(self._build_questionnaire_matrix_lines(rows))

        if table_kind == "narrative_matrix":
            return self._unique_lines(self._build_narrative_matrix_lines(rows))

        return self._unique_lines(self._build_generic_matrix_lines(rows))

    def _build_structured_questions(
        self,
        table_kind: str,
        rows: list[list[str]],
    ) -> list[ExtractedQuestion]:
        questions: list[ExtractedQuestion] = []

        if table_kind == "diagnostic_matrix":
            questions.extend(self._build_diagnostic_questions(rows))
        elif table_kind == "questionnaire_matrix":
            questions.extend(self._build_questionnaire_questions(rows))

        return self._unique_questions(questions)

    def _build_diagnostic_questions(self, rows: list[list[str]]) -> list[ExtractedQuestion]:
        questions: list[ExtractedQuestion] = []
        if len(rows) < 4:
            return questions

        header_row = rows[2]
        status_header = header_row[1].strip() if len(header_row) > 1 else "現況"

        for row in rows[3:]:
            section = row[0].strip() if row else ""
            if not section:
                continue

            status_value = row[1].strip() if len(row) > 1 else ""
            if not status_value:
                continue

            questions.extend(
                self._extract_checkbox_questions_from_text(
                    section=section,
                    field_label=status_header,
                    text=status_value,
                )
            )

        return questions

    def _build_questionnaire_questions(self, rows: list[list[str]]) -> list[ExtractedQuestion]:
        questions: list[ExtractedQuestion] = []
        if not rows:
            return questions

        header_rows = self._infer_header_rows("questionnaire_matrix", rows)
        header_depth = len(header_rows)
        headers = header_rows[-1] if header_rows else []
        data_rows = rows[header_depth:] if header_depth else rows

        for row_index, row in enumerate(data_rows, start=header_depth + 1):
            if not row or not any(cell.strip() for cell in row):
                continue

            row_label = self._to_inline_text(row[0]) or f"第{row_index}列"
            if headers:
                for column_index in range(1, min(len(row), len(headers))):
                    header = self._to_inline_text(headers[column_index]) or f"欄位{column_index + 1}"
                    value = row[column_index].strip()
                    if not value:
                        continue
                    questions.extend(
                        self._extract_checkbox_questions_from_text(
                            section=row_label,
                            field_label=f"{row_label} | {header}",
                            text=value,
                        )
                    )
                continue

            merged_value = self._merge_row_cells(row, start_column=1)
            if merged_value:
                questions.extend(
                    self._extract_checkbox_questions_from_text(
                        section=row_label,
                        field_label=row_label,
                        text=merged_value,
                    )
                )

        return questions

    def _extract_checkbox_questions_from_text(
        self,
        *,
        section: str,
        field_label: str,
        text: str,
    ) -> list[ExtractedQuestion]:
        questions: list[ExtractedQuestion] = []
        for index, (title, segment_text) in enumerate(
            self._split_question_segments(text, default_title=field_label),
            start=1,
        ):
            option_items = self._parse_checkbox_options(segment_text)
            if not option_items:
                continue

            options = [
                QuestionOption(
                    label=str(option["label"]),
                    selected=bool(option["selected"]),
                )
                for option in option_items
                if str(option.get("label", "")).strip()
            ]
            if not options:
                continue

            selected_values = [option.label for option in options if option.selected]
            question_title = self._normalize_question_title(title) or self._normalize_question_title(field_label)
            questions.append(
                ExtractedQuestion(
                    question_id=self._build_question_id(section, title, field_label, index),
                    question=question_title,
                    selection_type="multiple",
                    options=options,
                    selected_values=selected_values,
                )
            )

        return questions

    def _split_question_segments(
        self,
        text: str,
        default_title: str,
    ) -> list[tuple[str, str]]:
        segments: list[tuple[str, str]] = []
        current_title = default_title
        current_lines: list[str] = []

        def flush() -> None:
            nonlocal current_lines
            if current_lines and any(self.CHECKBOX_MARKER_PATTERN.search(line) for line in current_lines):
                segments.append((current_title, "\n".join(current_lines)))
            current_lines = []

        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line:
                continue

            if self._is_question_group_heading(line):
                flush()
                current_title = line
                continue

            current_lines.append(line)

        flush()

        if not segments and self._parse_checkbox_options(text):
            segments.append((default_title, text))
        return segments

    def _is_question_group_heading(self, line: str) -> bool:
        cleaned = self._to_inline_text(line)
        if not cleaned or self.CHECKBOX_MARKER_PATTERN.search(cleaned):
            return False
        if len(cleaned) > 70:
            return False
        if re.search(r"_{2,}", cleaned):
            return False
        if cleaned.startswith(("註", "例如", "說明", "A:", "B:", "C:", "A：", "B：", "C：")):
            return False

        heading_keywords = (
            "一般狀況",
            "孤島效應",
            "院內備援",
            "院外備援",
            "供水來源",
            "供水設備",
            "液氧槽規模",
            "液氧鋼瓶規模",
            "氧氣鋼瓶規模",
            "其他",
        )
        if any(keyword in cleaned for keyword in heading_keywords):
            return True

        if cleaned.endswith(("備援", "措施", "容量", "規模", "來源")):
            return True

        return False

    def _normalize_question_title(self, title: str) -> str:
        cleaned = self._to_inline_text(title)
        cleaned = re.sub(r"[（(]即不考慮孤島效應[）)]", "", cleaned)
        cleaned = re.sub(r"[（(]參考.*?[）)]", "", cleaned)
        cleaned = re.sub(r"^[\s|:：;；、,，]+", "", cleaned)
        cleaned = re.sub(r"[\s|:：;；、,，]+$", "", cleaned)
        return cleaned.strip()

    def _build_question_id(
        self,
        section: str,
        title: str,
        field_label: str,
        index: int,
    ) -> str:
        section_key = self._question_section_key(section)
        title_key = self._question_title_key(title) or self._question_title_key(field_label)

        if section_key == "it" and title_key == "internal_backup":
            return "internal_it_backup"
        if section_key == "it" and title_key == "external_backup":
            return "external_it_backup"

        if section_key and title_key:
            return f"{section_key}_{title_key}"
        if section_key:
            return f"{section_key}_question_{index}"
        if title_key:
            return title_key
        return f"question_{index}"

    def _question_section_key(self, section: str) -> str:
        if "基本資訊" in section:
            return "basic_info"
        if "供電" in section:
            return "power"
        if "供水" in section:
            return "water"
        if "醫用氣體" in section or "供氣" in section or "氧" in section:
            return "oxygen"
        if "資訊" in section or "IT" in section:
            return "it"
        if "設施維護" in section:
            return "facility"
        if "中央監控" in section:
            return "central_monitoring"
        return ""

    def _question_title_key(self, title: str) -> str:
        cleaned = self._normalize_question_title(title)
        if "院內備援" in cleaned:
            return "internal_backup"
        if "院外備援" in cleaned:
            return "external_backup"
        if "一般狀況" in cleaned:
            return "current_status"
        if "孤島效應" in cleaned:
            return "island_measures"
        if "供水來源" in cleaned:
            return "water_sources"
        if "供水設備" in cleaned:
            return "water_facilities"
        if "液氧槽" in cleaned:
            return "liquid_oxygen_tank"
        if "液氧鋼瓶" in cleaned:
            return "liquid_oxygen_cylinder"
        if "氧氣鋼瓶" in cleaned:
            return "oxygen_cylinder"
        if "管制設施" in cleaned:
            return "access_control"
        if "監控設施" in cleaned:
            return "monitoring_facilities"
        if "連通管道" in cleaned:
            return "pipeline_protection"
        if "可監視" in cleaned:
            return "operation_monitoring"
        ascii_key = re.sub(r"[^A-Za-z0-9]+", "_", cleaned).strip("_").lower()
        return ascii_key

    def _unique_questions(self, questions: list[ExtractedQuestion]) -> list[ExtractedQuestion]:
        counts: dict[str, int] = {}
        unique_questions: list[ExtractedQuestion] = []
        for question in questions:
            base_id = question.question_id or "question"
            counts[base_id] = counts.get(base_id, 0) + 1
            question_id = base_id if counts[base_id] == 1 else f"{base_id}_{counts[base_id]}"
            unique_questions.append(question.model_copy(update={"question_id": question_id}))
        return unique_questions

    def _build_table_warnings(
        self,
        table_kind: str,
        merged_ranges: list[MergedCellRange],
        semantic_lines: list[str],
    ) -> list[str]:
        warnings: list[str] = []
        if table_kind == "diagnostic_matrix":
            warnings.append("含不規則診斷表與合併儲存格，已額外展開 semantic_lines 供後續 LLM 判讀。")
        elif merged_ranges:
            warnings.append("表格含合併儲存格，已保留 cells、merged_ranges 與 semantic_lines 輔助判讀。")

        if not semantic_lines:
            warnings.append("表格未能產生額外語意 lines，後續需依 rows 與 cells 保守判讀。")
        return warnings

    def _build_metadata_block_lines(self, text: str) -> list[str]:
        lines: list[str] = []
        for raw_line in text.splitlines():
            cleaned = raw_line.strip()
            if not cleaned:
                continue

            label, value = self._split_label_value(cleaned)
            if label and value:
                lines.append(f"基本資料 | {label} = {self._to_inline_text(value)}")
            else:
                lines.append(f"基本資料 | {self._to_inline_text(cleaned)}")
        return lines

    def _build_diagnostic_matrix_lines(self, rows: list[list[str]]) -> list[str]:
        lines: list[str] = []
        if len(rows) < 4:
            return lines

        for metadata_row in rows[:2]:
            lines.extend(self._build_metadata_row_lines(metadata_row))

        header_row = rows[2]
        status_header = header_row[1].strip() or "現況"
        recommendation_header = self._merge_row_cells(header_row, start_column=2) or "建議"

        for row in rows[3:]:
            section = row[0].strip()
            if not section:
                continue

            status_value = row[1].strip() if len(row) > 1 else ""
            recommendation_value = self._merge_row_cells(row, start_column=2)

            if status_value:
                self._append_semantic_value(
                    lines=lines,
                    prefix=f"診斷表 | {section} | {status_header}",
                    value=status_value,
                    detail_label="現況細項",
                )

            if recommendation_value:
                self._append_semantic_value(
                    lines=lines,
                    prefix=f"診斷表 | {section} | {recommendation_header}",
                    value=recommendation_value,
                    detail_label="建議",
                )

        return lines

    def _build_questionnaire_matrix_lines(self, rows: list[list[str]]) -> list[str]:
        lines: list[str] = []
        if not rows:
            return lines

        header_rows = self._infer_header_rows("questionnaire_matrix", rows)
        header_depth = len(header_rows)
        headers = header_rows[-1] if header_rows else []
        data_rows = rows[header_depth:] if header_depth else rows

        for row_index, row in enumerate(data_rows, start=header_depth + 1):
            if not any(cell.strip() for cell in row):
                continue

            if len(row) == 1:
                lines.extend(self._build_metadata_block_lines(row[0]))
                continue

            row_label = self._to_inline_text(row[0]) or f"第{row_index}列"
            if headers:
                for column_index in range(1, min(len(row), len(headers))):
                    header = self._to_inline_text(headers[column_index]) or f"欄位{column_index + 1}"
                    value = row[column_index].strip()
                    if not value:
                        continue
                    self._append_semantic_value(
                        lines=lines,
                        prefix=f"問卷欄位 | {row_label} | {header}",
                        value=value,
                        detail_label="細項",
                    )
                continue

            merged_value = self._merge_row_cells(row, start_column=1)
            if merged_value:
                self._append_semantic_value(
                    lines=lines,
                    prefix=f"問卷欄位 | {row_label}",
                    value=merged_value,
                    detail_label="細項",
                )

        return lines

    def _build_generic_matrix_lines(self, rows: list[list[str]]) -> list[str]:
        lines: list[str] = []
        for row_index, row in enumerate(rows, start=1):
            non_empty_cells = [
                (column_index, value.strip())
                for column_index, value in enumerate(row, start=1)
                if value.strip()
            ]
            if not non_empty_cells:
                continue

            summary = "；".join(
                f"C{column_index}={self._to_inline_text(value)}"
                for column_index, value in non_empty_cells[:6]
            )
            lines.append(f"表格列{row_index} | {summary}")

            for column_index, value in non_empty_cells[:6]:
                self._append_semantic_value(
                    lines=lines,
                    prefix=f"表格列{row_index} | 欄位{column_index}",
                    value=value,
                    detail_label="細項",
                )

        return lines

    def _build_narrative_matrix_lines(self, rows: list[list[str]]) -> list[str]:
        lines: list[str] = []
        for row_index, row in enumerate(rows, start=1):
            if not any(cell.strip() for cell in row):
                continue

            left_value = row[0].strip() if row else ""
            right_value = row[1].strip() if len(row) > 1 else ""

            if left_value:
                self._append_semantic_value(
                    lines=lines,
                    prefix=f"問卷敘述 | 第{row_index}列左欄",
                    value=left_value,
                    detail_label="主敘述",
                )

            if right_value:
                self._append_semantic_value(
                    lines=lines,
                    prefix=f"問卷敘述 | 第{row_index}列右欄",
                    value=right_value,
                    detail_label="註記",
                )

        return lines

    def _build_metadata_row_lines(self, row: list[str]) -> list[str]:
        lines: list[str] = []
        for index in range(0, len(row), 2):
            label = row[index].strip()
            value = row[index + 1].strip() if index + 1 < len(row) else ""
            if label and value:
                lines.append(f"基本資料 | {label} = {self._to_inline_text(value)}")
        return lines

    def _append_semantic_value(
        self,
        lines: list[str],
        prefix: str,
        value: str,
        detail_label: str,
    ) -> None:
        inline_value = self._to_inline_text(value)
        if inline_value:
            lines.append(f"{prefix} = {inline_value}")

        checkbox_options = self._parse_checkbox_options(value)
        if checkbox_options:
            lines.extend(self._build_checkbox_semantic_lines(prefix, checkbox_options))

        chunks = self._split_semantic_chunks(value)
        if len(chunks) <= 1:
            return

        for index, chunk in enumerate(chunks, start=1):
            chunk_text = self._to_inline_text(chunk)
            if not chunk_text:
                continue
            lines.append(f"{prefix} | {detail_label}{index} = {chunk_text}")

    def _split_semantic_chunks(self, value: str) -> list[str]:
        chunks: list[str] = []
        for raw_line in value.splitlines():
            cleaned = raw_line.strip()
            if not cleaned:
                continue

            checkbox_parts = self._split_checkbox_chunks(cleaned)
            if len(checkbox_parts) > 1:
                chunks.extend(checkbox_parts)
            else:
                chunks.append(cleaned)

        if not chunks:
            inline_value = self._to_inline_text(value)
            if inline_value:
                chunks.append(inline_value)

        return chunks[:20]

    def _split_checkbox_chunks(self, value: str) -> list[str]:
        if len(self.CHECKBOX_MARKER_PATTERN.findall(value)) < 2:
            return [value]

        parts = [
            part.strip()
            for part in self.CHECKBOX_MARKER_PATTERN.split(value)
            if part.strip()
        ]
        merged_parts: list[str] = []
        index = 0
        while index < len(parts):
            current = parts[index]
            if self._is_checkbox_symbol(current) and index + 1 < len(parts):
                merged_parts.append(f"{current}{parts[index + 1]}".strip())
                index += 2
                continue
            merged_parts.append(current)
            index += 1

        parts = merged_parts
        return parts or [value]

    def _parse_checkbox_options(self, text: str) -> list[dict[str, object]]:
        matches = list(self.CHECKBOX_MARKER_PATTERN.finditer(text))
        if not matches:
            return []

        options: list[dict[str, object]] = []
        for index, match in enumerate(matches):
            symbol = match.group(1)
            label_start = match.end()
            label_end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
            label = self._clean_checkbox_label(text[label_start:label_end])
            if not label:
                continue

            options.append(
                {
                    "label": label,
                    "selected": self._is_checked_symbol(symbol),
                    "symbol": symbol,
                }
            )

        return options

    def _build_checkbox_semantic_lines(
        self,
        prefix: str,
        options: list[dict[str, object]],
    ) -> list[str]:
        if not options:
            return []

        option_parts: list[str] = []
        selected_values: list[str] = []
        for option in options[:30]:
            label = str(option["label"])
            selected = bool(option["selected"])
            state = "selected" if selected else "not_selected"
            option_parts.append(f"{label}: {state}")
            if selected:
                selected_values.append(label)

        lines = [
            f"{prefix} | selection_type = multiple",
            f"{prefix} | checkbox_options = {'; '.join(option_parts)}",
        ]
        if selected_values:
            lines.append(f"{prefix} | selected_values = {'; '.join(selected_values)}")
        return lines

    def _clean_checkbox_label(self, label: str) -> str:
        cleaned = self._to_inline_text(label)
        for context_marker in (
            "；孤島效應下",
            "；院外備援",
            "；一般狀況",
            "；申請建置項目",
            "；用途",
            "；蓄水量",
            "；可累積供水",
            "；如機房",
            "；如可監視",
            "；註:",
            "；註：",
        ):
            if context_marker in cleaned:
                cleaned = cleaned.split(context_marker, 1)[0]
        cleaned = re.sub(r"^[\s:：;；、,，.。)）(（]+", "", cleaned)
        cleaned = re.sub(r"[\s:：;；、,，]+$", "", cleaned)
        return cleaned.strip()

    def _is_checkbox_symbol(self, symbol: str) -> bool:
        return symbol in self.CHECKED_SYMBOLS or symbol in self.UNCHECKED_SYMBOLS or symbol in self.TEXT_CHECKED_SYMBOLS

    def _is_checked_symbol(self, symbol: str) -> bool:
        return symbol in self.CHECKED_SYMBOLS or symbol in self.TEXT_CHECKED_SYMBOLS

    def _split_label_value(self, value: str) -> tuple[str, str]:
        for separator in ("：", ":"):
            if separator in value:
                label, content = value.split(separator, 1)
                return label.strip(), content.strip()
        return "", ""

    def _merge_row_cells(self, row: list[str], start_column: int) -> str:
        return "\n".join(cell.strip() for cell in row[start_column:] if cell.strip())

    def _looks_like_header_row(self, row: list[str]) -> bool:
        non_empty_cells = [cell.strip() for cell in row if cell.strip()]
        if len(non_empty_cells) < 2:
            return False

        if any("\n" in cell for cell in non_empty_cells):
            return False

        if max(len(cell) for cell in non_empty_cells) > 40:
            return False

        header_keywords = (
            "項目",
            "一般狀況",
            "現況",
            "應變措施",
            "建議",
            "類別",
            "補助前",
            "補助後",
            "天數",
            "達標",
            "效益",
            "備註",
        )
        return any(keyword in cell for keyword in header_keywords for cell in non_empty_cells)

    def _to_inline_text(self, value: str) -> str:
        return "；".join(part.strip() for part in value.splitlines() if part.strip())

    def _unique_lines(self, lines: list[str]) -> list[str]:
        seen: set[str] = set()
        unique_lines: list[str] = []
        for line in lines:
            cleaned = line.strip()
            if not cleaned or cleaned in seen:
                continue
            seen.add(cleaned)
            unique_lines.append(cleaned)
        return unique_lines

    def _extract_table_grid(
        self,
        table: Table,
    ) -> tuple[list[list[str]], list[TableCell], int, int]:
        grid_rows: list[list[str]] = []
        extracted_cells: list[TableCell] = []
        active_vertical_merges: dict[int, TableCell] = {}
        column_count = 0

        for row_index, row in enumerate(table.rows, start=1):
            logical_row = [""] * row.grid_cols_before
            logical_column = row.grid_cols_before + 1
            next_active_vertical_merges: dict[int, TableCell] = {}

            for cell_element in row._tr.tc_lst:
                column_span = self._get_column_span(cell_element)
                cell_text = self._extract_cell_text(cell_element)
                vertical_merge_state = self._get_vertical_merge_state(cell_element)
                self._ensure_row_width(logical_row, logical_column - 1 + column_span)

                if vertical_merge_state == "continue":
                    origin_cell = active_vertical_merges.get(logical_column)
                    if origin_cell is not None:
                        origin_cell.row_span += 1
                        for offset in range(column_span):
                            next_active_vertical_merges[logical_column + offset] = origin_cell
                        logical_column += column_span
                        continue

                extracted_cell = TableCell(
                    row_index=row_index,
                    column_index=logical_column,
                    text=cell_text,
                    row_span=1,
                    column_span=column_span,
                )
                extracted_cells.append(extracted_cell)
                logical_row[logical_column - 1] = cell_text

                for offset in range(1, column_span):
                    logical_row[logical_column - 1 + offset] = ""

                if vertical_merge_state == "restart":
                    for offset in range(column_span):
                        next_active_vertical_merges[logical_column + offset] = extracted_cell

                logical_column += column_span

            logical_row.extend([""] * row.grid_cols_after)
            column_count = max(column_count, len(logical_row))
            grid_rows.append(logical_row)
            active_vertical_merges = next_active_vertical_merges

        return grid_rows, extracted_cells, len(grid_rows), column_count

    def _extract_paragraph_text(self, paragraph: Paragraph) -> str:
        return self._extract_text_from_element(paragraph._element)

    def _extract_cell_text(self, cell_element: CT_Tc) -> str:
        paragraph_texts: list[str] = []
        for child in cell_element.iterchildren():
            if child.tag == qn("w:p"):
                text = self._extract_text_from_element(child)
                if text:
                    paragraph_texts.append(text)
        return "\n".join(paragraph_texts)

    def _extract_text_from_element(self, element: object) -> str:
        parts: list[str] = []
        for node in element.iter():
            if node.tag == qn("w:t"):
                parts.append(self._normalize_inline_symbols(node.text or ""))
            elif node.tag == qn("w:tab"):
                parts.append("\t")
            elif node.tag in {qn("w:br"), qn("w:cr")}:
                parts.append("\n")
            elif node.tag == qn("w:noBreakHyphen"):
                parts.append("-")
            elif node.tag == qn("w:sym"):
                parts.append(self._normalize_symbol(node))
        return self._normalize_text("".join(parts))

    def _normalize_symbol(self, symbol_element: object) -> str:
        font_name = (symbol_element.get(qn("w:font")) or "").strip().lower()
        char_code = (symbol_element.get(qn("w:char")) or "").strip().upper()
        if not char_code:
            return ""

        direct_match = self.SYMBOL_MAP.get((font_name, char_code))
        if direct_match is not None:
            return direct_match

        if char_code in self.DEFAULT_SYMBOL_MAP:
            return self.DEFAULT_SYMBOL_MAP[char_code]

        if font_name.startswith("wingdings"):
            return "○"

        return ""

    def _get_column_span(self, cell_element: CT_Tc) -> int:
        tc_properties = cell_element.tcPr
        if tc_properties is None or tc_properties.gridSpan is None:
            return 1
        try:
            return int(tc_properties.gridSpan.val)
        except (TypeError, ValueError):
            return 1

    def _get_vertical_merge_state(self, cell_element: CT_Tc) -> str | None:
        tc_properties = cell_element.tcPr
        if tc_properties is None or tc_properties.vMerge is None:
            return None
        if tc_properties.vMerge.val == "restart":
            return "restart"
        return "continue"

    def _rows_to_markdown(self, rows: list[list[str]]) -> str:
        if not rows:
            return ""

        width = max(len(row) for row in rows)
        normalized_rows = [row + [""] * (width - len(row)) for row in rows]
        header = normalized_rows[0]
        divider = ["---"] * width
        markdown_lines = [
            "| " + " | ".join(self._escape_markdown_cell(cell) for cell in header) + " |",
            "| " + " | ".join(divider) + " |",
        ]
        for row in normalized_rows[1:]:
            markdown_lines.append(
                "| " + " | ".join(self._escape_markdown_cell(cell) for cell in row) + " |"
            )
        return "\n".join(markdown_lines)

    def _escape_markdown_cell(self, value: str) -> str:
        return value.replace("|", "\\|").replace("\n", "<br>")

    def _normalize_text(self, value: str | None) -> str:
        if not value:
            return ""

        lines = []
        for raw_line in value.replace("\xa0", " ").replace("\u3000", " ").splitlines():
            normalized_line = re.sub(r"[ \t\r\f\v]+", " ", raw_line).strip()
            if normalized_line:
                lines.append(normalized_line)
        return "\n".join(lines)

    def _ensure_row_width(self, row: list[str], width: int) -> None:
        if len(row) < width:
            row.extend([""] * (width - len(row)))

    def _normalize_inline_symbols(self, text: str) -> str:
        normalized_text = text
        for source_symbol, target_symbol in self.TEXT_SYMBOL_MAP.items():
            normalized_text = normalized_text.replace(source_symbol, target_symbol)
        return normalized_text
