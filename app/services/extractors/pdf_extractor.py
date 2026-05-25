from __future__ import annotations

from pathlib import Path

import pdfplumber

from app.models.schemas import ExtractedSource, ExtractedTable


class PdfExtractor:
    def extract(self, file_path: Path) -> ExtractedSource:
        try:
            with pdfplumber.open(file_path) as pdf:
                page_texts: list[str] = []
                tables: list[ExtractedTable] = []

                for page_number, page in enumerate(pdf.pages, start=1):
                    text = self._clean_text(page.extract_text())
                    if text:
                        page_texts.append(f"[Page {page_number}]\n{text}")

                    raw_tables = page.extract_tables() or []
                    for table_index, raw_table in enumerate(raw_tables, start=1):
                        rows = self._normalize_rows(raw_table)
                        if not rows:
                            continue

                        tables.append(
                            ExtractedTable(
                                title=f"Page {page_number} Table {table_index}",
                                page_number=page_number,
                                columns=rows[0],
                                rows=rows,
                                markdown=self._rows_to_markdown(rows),
                                row_count=len(rows),
                                column_count=max(len(row) for row in rows),
                            )
                        )
        except Exception as exc:
            raise ValueError(f"無法讀取 PDF 檔案: {file_path.name}") from exc

        text = "\n\n".join(page_texts).strip()
        warnings: list[str] = []
        if len(text) < 50:
            warnings.append("PDF 抽取文字量偏少，可能是掃描型 PDF，後續可能需要影像辨識。")

        if not text and not tables:
            warnings.append("PDF 幾乎無可抽取內容。")

        return ExtractedSource(
            type="pdf",
            source_name=file_path.name,
            text=text or None,
            tables=tables,
            warnings=warnings,
            metadata={"path": str(file_path)},
        )

    def _normalize_rows(self, raw_table: list[list[str | None]] | None) -> list[list[str]]:
        if not raw_table:
            return []

        rows: list[list[str]] = []
        for row in raw_table:
            cleaned_row = [self._clean_text(cell) for cell in row]
            if any(cell for cell in cleaned_row):
                rows.append(cleaned_row)
        return rows

    def _rows_to_markdown(self, rows: list[list[str]]) -> str:
        if not rows:
            return ""

        width = max(len(row) for row in rows)
        normalized_rows = [row + [""] * (width - len(row)) for row in rows]
        header = normalized_rows[0]
        divider = ["---"] * width
        markdown_lines = [
            "| " + " | ".join(header) + " |",
            "| " + " | ".join(divider) + " |",
        ]
        for row in normalized_rows[1:]:
            markdown_lines.append("| " + " | ".join(row) + " |")
        return "\n".join(markdown_lines)

    def _clean_text(self, value: str | None) -> str:
        if not value:
            return ""
        return " ".join(value.split()).strip()
