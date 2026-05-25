from __future__ import annotations

import json
from functools import cached_property
from typing import Any

from openai import OpenAI
from pydantic import ValidationError

from app.config import get_settings
from app.models.schemas import EvidenceText, KeyFinding, NormalizedDocument, ReportJSON
from app.prompts.report_prompts import (
    build_report_expansion_user_prompt,
    build_report_user_prompt,
    ensure_template_supported,
    get_report_expansion_system_prompt,
    get_report_system_prompt,
)


class ReportPlanner:
    MAX_PLANNING_PAYLOAD_CHARS = 45000
    MAX_SOURCE_PARAGRAPHS = 20
    MAX_SOURCE_TABLES = 8
    MAX_TABLE_ROWS = 18
    MAX_TABLE_SEMANTIC_LINES = 70
    MAX_FACTS = 140
    MAX_FACT_CHARS = 1400
    IMPORTANT_FACT_KEYWORDS = (
        "selected_values",
        "checkbox_options",
        "建議",
        "診斷",
        "總體觀察",
        "供電",
        "發電",
        "油",
        "UPS",
        "供水",
        "儲水",
        "病床",
        "供氣",
        "氧",
        "醫用氣體",
        "資訊",
        "IT",
        "備援",
        "中央監控",
        "工安",
        "孤島",
        "72",
        "天",
        "小時",
        "補助",
        "效益",
    )

    def __init__(self, model: str | None = None) -> None:
        self.settings = get_settings()
        self.model = model or self.settings.openai_model

    @cached_property
    def client(self) -> OpenAI:
        if not self.settings.openai_api_key:
            raise RuntimeError("OPENAI_API_KEY 尚未設定。")
        return OpenAI(api_key=self.settings.openai_api_key)

    def generate_report(self, normalized_document: NormalizedDocument) -> ReportJSON:
        ensure_template_supported(normalized_document.template_id)

        normalized_payload = self._build_planning_payload(normalized_document)

        parsed = self._cleanup_report(
            self._request_report_json(
            instructions=get_report_system_prompt(normalized_document.template_id),
            prompt=build_report_user_prompt(normalized_payload),
            max_output_tokens=4500,
            )
        )

        if self._is_sparse_report(parsed):
            expanded = self._try_expand_report(
                normalized_document=normalized_document,
                normalized_payload=normalized_payload,
                current_report=parsed,
            )
            if expanded is not None:
                return self._cleanup_report(expanded)

        return parsed

    def _build_planning_payload(self, normalized_document: NormalizedDocument) -> str:
        """Build a compact payload for the LLM instead of sending raw attachments."""
        payload: dict[str, Any] = {
            "project_name": normalized_document.project_name,
            "hospital_name": normalized_document.hospital_name,
            "template_id": normalized_document.template_id,
            "warnings": normalized_document.warnings[:30],
            "sources": [
                self._compact_source(source)
                for source in normalized_document.sources
            ],
            "facts": self._select_planning_facts(normalized_document),
        }

        text = self._dump_json(payload)
        if len(text) > self.MAX_PLANNING_PAYLOAD_CHARS:
            self._shrink_sources_for_context(payload)
            text = self._dump_json(payload)

        while len(text) > self.MAX_PLANNING_PAYLOAD_CHARS and payload["facts"]:
            payload["facts"].pop()
            text = self._dump_json(payload)

        if len(text) > self.MAX_PLANNING_PAYLOAD_CHARS:
            for source in payload["sources"]:
                source["paragraphs"] = source.get("paragraphs", [])[:5]
                if "text_excerpt" in source:
                    source["text_excerpt"] = self._truncate_text(source["text_excerpt"], 1000)
            text = self._dump_json(payload)

        return text[: self.MAX_PLANNING_PAYLOAD_CHARS]

    def _shrink_sources_for_context(self, payload: dict[str, Any]) -> None:
        for source in payload["sources"]:
            source["tables"] = source.get("tables", [])[:3]
            for table in source["tables"]:
                table["rows"] = table.get("rows", [])[:6]
                table["raw_rows"] = table.get("raw_rows", [])[:6]
                table["semantic_lines"] = table.get("semantic_lines", [])[:30]

    def _compact_source(self, source: Any) -> dict[str, Any]:
        compact: dict[str, Any] = {
            "type": source.type,
            "source_name": source.source_name,
            "warnings": source.warnings[:10],
            "metadata": self._compact_metadata(source.metadata),
        }

        if source.paragraphs:
            compact["paragraphs"] = [
                self._truncate_text(paragraph, 500)
                for paragraph in source.paragraphs[: self.MAX_SOURCE_PARAGRAPHS]
            ]

        if source.text:
            compact["text_excerpt"] = self._truncate_text(source.text, 2500)

        if source.tables:
            compact["tables"] = [
                self._compact_table(table)
                for table in source.tables[: self.MAX_SOURCE_TABLES]
            ]

        if source.image_analysis:
            compact["image_analysis"] = source.image_analysis.model_dump(mode="json")

        return compact

    def _compact_metadata(self, metadata: dict[str, Any]) -> dict[str, Any]:
        skipped_keys = {"data_url", "base64", "raw_xml"}
        compact: dict[str, Any] = {}
        for key, value in metadata.items():
            if key in skipped_keys:
                continue
            compact[key] = self._truncate_value(value, max_chars=800)
        return compact

    def _compact_table(self, table: Any) -> dict[str, Any]:
        return {
            "title": table.title,
            "sheet_name": table.sheet_name,
            "table_kind": table.table_kind,
            "columns": table.columns[:20],
            "row_count": table.row_count,
            "column_count": table.column_count,
            "header_rows": table.header_rows[:4],
            "semantic_lines": self._select_semantic_lines(table.semantic_lines),
            "questions": [
                question.model_dump(mode="json")
                for question in table.questions[:80]
            ],
            "rows": self._truncate_value(table.rows[: self.MAX_TABLE_ROWS], max_chars=5000),
            "raw_rows": self._truncate_value(table.raw_rows[: self.MAX_TABLE_ROWS], max_chars=5000),
            "warnings": table.warnings[:10],
        }

    def _select_semantic_lines(self, semantic_lines: list[str]) -> list[str]:
        priority: list[str] = []
        secondary: list[str] = []
        seen: set[str] = set()

        for line in semantic_lines:
            if line in seen:
                continue
            seen.add(line)
            if self._is_important_text(line):
                priority.append(self._truncate_text(line, 900))
            elif len(secondary) < self.MAX_TABLE_SEMANTIC_LINES:
                secondary.append(self._truncate_text(line, 500))

        selected = priority + secondary
        return selected[: self.MAX_TABLE_SEMANTIC_LINES]

    def _select_planning_facts(
        self,
        normalized_document: NormalizedDocument,
    ) -> list[dict[str, Any]]:
        scored_facts: list[tuple[int, int, Any]] = []
        for index, fact in enumerate(normalized_document.facts):
            score = self._score_fact(fact.content, fact.fact_type)
            scored_facts.append((score, index, fact))

        selected = [
            (index, fact)
            for score, index, fact in sorted(scored_facts, key=lambda item: (-item[0], item[1]))
            if score > 0
        ][: self.MAX_FACTS]

        if len(selected) < self.MAX_FACTS:
            selected_indexes = {index for index, _ in selected}
            for _, index, fact in scored_facts:
                if index in selected_indexes:
                    continue
                selected.append((index, fact))
                if len(selected) >= self.MAX_FACTS:
                    break

        return [
            {
                "source_name": fact.source_name,
                "fact_type": fact.fact_type,
                "block_order": fact.block_order,
                "content": self._truncate_text(fact.content, self.MAX_FACT_CHARS),
            }
            for index, fact in sorted(selected, key=lambda item: item[0])
        ]

    def _score_fact(self, content: str, fact_type: str) -> int:
        score = 0
        if fact_type == "table_semantic":
            score += 35
        elif fact_type == "table_questions":
            score += 45
        elif fact_type == "table_structure":
            score += 20
        elif fact_type in {"image_analysis", "image_checkbox"}:
            score += 15

        for keyword in self.IMPORTANT_FACT_KEYWORDS:
            if keyword in content:
                score += 8

        if "selected_values" in content:
            score += 40
        if "checkbox_options" in content:
            score += 30
        if "selection_type=multiple" in content:
            score += 20
        return score

    def _is_important_text(self, text: str) -> bool:
        return any(keyword in text for keyword in self.IMPORTANT_FACT_KEYWORDS)

    def _truncate_value(self, value: Any, max_chars: int) -> Any:
        if isinstance(value, str):
            return self._truncate_text(value, max_chars)
        if isinstance(value, list):
            truncated: list[Any] = []
            used = 0
            for item in value:
                compact_item = self._truncate_value(item, max_chars=max_chars)
                item_size = len(json.dumps(compact_item, ensure_ascii=False, default=str))
                if used + item_size > max_chars and truncated:
                    break
                truncated.append(compact_item)
                used += item_size
            return truncated
        if isinstance(value, dict):
            return {
                str(key): self._truncate_value(item, max_chars=max_chars // 2)
                for key, item in value.items()
            }
        return value

    def _truncate_text(self, text: str, max_chars: int) -> str:
        normalized = " ".join(str(text).split())
        if len(normalized) <= max_chars:
            return normalized
        return f"{normalized[: max_chars - 20]}... [truncated]"

    def _dump_json(self, payload: dict[str, Any]) -> str:
        return json.dumps(payload, ensure_ascii=False, indent=2, default=str)

    def _try_expand_report(
        self,
        normalized_document: NormalizedDocument,
        normalized_payload: str,
        current_report: ReportJSON,
    ) -> ReportJSON | None:
        current_report_payload = json.dumps(
            current_report.model_dump(mode="json"),
            ensure_ascii=False,
            indent=2,
        )

        try:
            expanded = self._request_report_json(
                instructions=get_report_expansion_system_prompt(
                    normalized_document.template_id,
                ),
                prompt=build_report_expansion_user_prompt(
                    normalized_document_json=normalized_payload,
                    current_report_json=current_report_payload,
                ),
                max_output_tokens=5500,
            )
        except RuntimeError:
            return None

        if self._report_density_score(expanded) <= self._report_density_score(current_report):
            return current_report
        return expanded

    def _cleanup_report(self, report: ReportJSON) -> ReportJSON:
        intro = self._clean_paragraph_list(report.introduction_paragraphs)
        diagnosis = self._clean_evidenced_text_list(
            report.diagnosis_paragraphs,
            require_evidence=True,
        )
        recommendation_sections = []
        for section in report.recommendation_sections:
            cleaned_title = section.title.strip()
            cleaned_paragraphs = self._clean_evidenced_text_list(
                section.paragraphs,
                forbidden_exact={cleaned_title},
                require_evidence=True,
            )
            if not cleaned_title or not cleaned_paragraphs:
                continue
            recommendation_sections.append(
                section.model_copy(
                    update={
                        "title": cleaned_title,
                        "paragraphs": cleaned_paragraphs,
                    }
                )
            )

        recommendations = self._clean_evidenced_text_list(
            report.recommendations,
            require_evidence=True,
        )
        missing_information = self._clean_paragraph_list(report.missing_information)
        key_findings = self._clean_key_findings(report.key_findings)

        return report.model_copy(
            update={
                "key_findings": key_findings,
                "introduction_paragraphs": intro,
                "diagnosis_paragraphs": diagnosis,
                "recommendation_sections": recommendation_sections,
                "recommendations": recommendations,
                "missing_information": missing_information,
                "recommendation_intro": report.recommendation_intro.strip(),
                "summary": report.summary.strip(),
                "background": report.background.strip(),
                "title": report.title.strip(),
            }
        )

    def _request_report_json(
        self,
        instructions: str,
        prompt: str,
        max_output_tokens: int,
    ) -> ReportJSON:
        attempts = [
            {
                "mode": "structured",
                "temperature": 0.1,
                "max_output_tokens": max_output_tokens,
                "prompt": prompt,
            },
            {
                "mode": "structured",
                "temperature": 0,
                "max_output_tokens": max(max_output_tokens + 1500, 6500),
                "prompt": self._build_retry_prompt(prompt, compact=False),
            },
            {
                "mode": "raw",
                "temperature": 0,
                "max_output_tokens": max(max_output_tokens + 1500, 6500),
                "prompt": self._build_retry_prompt(prompt, compact=True),
            },
        ]

        last_error: Exception | None = None
        for attempt in attempts:
            try:
                if attempt["mode"] == "structured":
                    return self._request_structured_report_json(
                        instructions=instructions,
                        prompt=attempt["prompt"],
                        temperature=attempt["temperature"],
                        max_output_tokens=attempt["max_output_tokens"],
                    )
                return self._request_raw_report_json(
                    instructions=instructions,
                    prompt=attempt["prompt"],
                    temperature=attempt["temperature"],
                    max_output_tokens=attempt["max_output_tokens"],
                )
            except Exception as exc:
                last_error = exc

        raise RuntimeError(
            f"OpenAI 報告規劃呼叫失敗: {self._summarize_openai_error(last_error)}"
        ) from last_error

    def _request_structured_report_json(
        self,
        instructions: str,
        prompt: str,
        temperature: float,
        max_output_tokens: int,
    ) -> ReportJSON:
        response = self.client.responses.parse(
            model=self.model,
            instructions=instructions,
            input=self._build_report_input(prompt),
            text_format=ReportJSON,
            temperature=temperature,
            max_output_tokens=max_output_tokens,
        )

        parsed = response.output_parsed
        if parsed is None:
            raise RuntimeError("OpenAI 沒有回傳可解析的 report_json。")
        return parsed

    def _request_raw_report_json(
        self,
        instructions: str,
        prompt: str,
        temperature: float,
        max_output_tokens: int,
    ) -> ReportJSON:
        response = self.client.responses.create(
            model=self.model,
            instructions=instructions,
            input=self._build_report_input(prompt),
            temperature=temperature,
            max_output_tokens=max_output_tokens,
        )

        output_text = (response.output_text or "").strip()
        if not output_text:
            raise RuntimeError("OpenAI 沒有回傳任何文字內容。")
        return self._parse_report_json_text(output_text)

    def _build_report_input(self, prompt: str) -> list[dict[str, object]]:
        return [
            {
                "role": "user",
                "content": [
                    {
                        "type": "input_text",
                        "text": prompt,
                    }
                ],
            }
        ]

    def _build_retry_prompt(self, prompt: str, compact: bool) -> str:
        retry_note = [
            "",
            "重要：你必須只輸出單一完整 JSON 物件，不可輸出 markdown code fence、說明文字或註解。",
            "所有欄位都必須符合 schema，且 JSON 必須完整結束，不可截斷。",
        ]
        if compact:
            retry_note.append(
                "若內容過長，請縮短每段文字，但不要省略必要欄位與結構。"
            )
        return f"{prompt}\n\n" + "\n".join(retry_note)

    def _parse_report_json_text(self, text: str) -> ReportJSON:
        candidate_texts = [text, self._strip_code_fences(text)]
        extracted_json = self._extract_json_object(text)
        if extracted_json:
            candidate_texts.append(extracted_json)

        last_error: Exception | None = None
        seen: set[str] = set()
        for candidate in candidate_texts:
            cleaned = candidate.strip()
            if not cleaned or cleaned in seen:
                continue
            seen.add(cleaned)
            try:
                return ReportJSON.model_validate_json(cleaned)
            except ValidationError as exc:
                last_error = exc

        if last_error is None:
            raise RuntimeError("OpenAI 回傳內容為空，無法解析 report_json。")
        raise last_error

    def _strip_code_fences(self, text: str) -> str:
        stripped = text.strip()
        if not stripped.startswith("```"):
            return stripped

        lines = stripped.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].startswith("```"):
            lines = lines[:-1]
        return "\n".join(lines).strip()

    def _extract_json_object(self, text: str) -> str | None:
        start = text.find("{")
        end = text.rfind("}")
        if start == -1 or end == -1 or end <= start:
            return None
        return text[start : end + 1]

    def _summarize_openai_error(self, error: Exception | None) -> str:
        if error is None:
            return "未知錯誤"
        message = " ".join(str(error).split())
        if len(message) > 240:
            return f"{message[:237]}..."
        return message

    def _is_sparse_report(self, report: ReportJSON) -> bool:
        if len(report.introduction_paragraphs) < 2:
            return True
        if len(report.diagnosis_paragraphs) < 4:
            return True
        if len(report.key_findings) < 4:
            return True
        if len(report.recommendation_sections) < 3:
            return True

        total_recommendation_paragraphs = sum(
            len(section.paragraphs)
            for section in report.recommendation_sections
        )
        if total_recommendation_paragraphs < 8:
            return True

        if any(len(section.paragraphs) == 0 for section in report.recommendation_sections):
            return True

        grounded_paragraph_count = sum(
            1
            for section in report.recommendation_sections
            for paragraph in section.paragraphs
            if self._has_evidence(paragraph)
        )
        return grounded_paragraph_count < 3

    def _report_density_score(self, report: ReportJSON) -> int:
        return (
            len(report.introduction_paragraphs) * 2
            + len(report.diagnosis_paragraphs) * 4
            + len(report.key_findings) * 3
            + len(report.recommendation_sections) * 3
            + sum(len(section.paragraphs) * 2 for section in report.recommendation_sections)
            + len(report.recommendations)
        )

    def _clean_paragraph_list(
        self,
        paragraphs: list[str],
        forbidden_exact: set[str] | None = None,
    ) -> list[str]:
        cleaned: list[str] = []
        seen: set[str] = set()
        forbidden = forbidden_exact or set()
        for paragraph in paragraphs:
            text = paragraph.strip()
            if not text:
                continue
            if text in forbidden:
                continue
            if self._is_heading_like(text):
                continue
            if text in seen:
                continue
            cleaned.append(text)
            seen.add(text)
        return cleaned

    def _clean_evidenced_text_list(
        self,
        paragraphs: list[EvidenceText],
        forbidden_exact: set[str] | None = None,
        require_evidence: bool = False,
    ) -> list[EvidenceText]:
        cleaned: list[EvidenceText] = []
        seen: set[str] = set()
        forbidden = forbidden_exact or set()
        for paragraph in paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue
            if text in forbidden:
                continue
            if self._is_heading_like(text):
                continue
            if require_evidence and not self._has_evidence(paragraph):
                continue
            if text in seen:
                continue
            cleaned.append(paragraph.model_copy(update={"text": text}))
            seen.add(text)
        return cleaned

    def _clean_key_findings(self, findings: list[KeyFinding]) -> list[KeyFinding]:
        cleaned: list[KeyFinding] = []
        seen: set[tuple[str, str]] = set()
        for finding in findings:
            item = finding.item.strip()
            description = finding.description.strip()
            evidence = [
                item
                for item in finding.evidence
                if item.source_name.strip() or item.content.strip()
            ]
            if not item or not description or not evidence:
                continue
            key = (item, description)
            if key in seen:
                continue
            cleaned.append(
                finding.model_copy(
                    update={
                        "item": item,
                        "description": description,
                        "evidence": evidence,
                    }
                )
            )
            seen.add(key)
        return cleaned

    def _has_evidence(self, paragraph: EvidenceText) -> bool:
        return any(
            evidence.source_name.strip() or evidence.content.strip()
            for evidence in paragraph.evidence
        )

    def _is_heading_like(self, text: str) -> bool:
        return text.startswith(("壹、", "貳、", "參、", "肆、"))
