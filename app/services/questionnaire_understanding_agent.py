from __future__ import annotations

import json
import re
from functools import cached_property

from openai import OpenAI

from app.config import get_settings
from app.models.schemas import HospitalResilienceDiagnosisForm, NormalizedDocument
from app.prompts.questionnaire_prompts import (
    QUESTIONNAIRE_UNDERSTANDING_SYSTEM_PROMPT,
    build_questionnaire_understanding_user_prompt,
)


class QuestionnaireUnderstandingAgent:
    def __init__(self, model: str | None = None) -> None:
        self.settings = get_settings()
        self.model = model or self.settings.openai_model

    @cached_property
    def client(self) -> OpenAI:
        if not self.settings.openai_api_key:
            raise RuntimeError("OPENAI_API_KEY 未設定。")
        return OpenAI(api_key=self.settings.openai_api_key)

    def analyze(
        self, normalized_document: NormalizedDocument
    ) -> HospitalResilienceDiagnosisForm:
        normalized_payload = json.dumps(
            normalized_document.model_dump(mode="json"),
            ensure_ascii=False,
            indent=2,
        )

        try:
            response = self.client.responses.create(
                model=self.model,
                instructions=QUESTIONNAIRE_UNDERSTANDING_SYSTEM_PROMPT,
                input=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "input_text",
                                "text": build_questionnaire_understanding_user_prompt(
                                    normalized_payload
                                ),
                            }
                        ],
                    }
                ],
                temperature=0,
                max_output_tokens=5000,
            )
        except Exception as exc:
            raise RuntimeError(f"OpenAI 產生問卷理解結果失敗: {exc}") from exc

        output_text = getattr(response, "output_text", "") or ""
        if not output_text.strip():
            raise RuntimeError("OpenAI 沒有回傳可解析的 understanding_json。")

        json_text = _extract_json_object(output_text)

        try:
            payload = json.loads(json_text)
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                f"OpenAI 回傳內容不是合法 JSON: {exc}. 原始輸出片段: {output_text[:500]}"
            ) from exc

        try:
            return HospitalResilienceDiagnosisForm.model_validate(payload)
        except Exception as exc:
            raise RuntimeError(
                f"OpenAI 回傳 JSON 結構不符合預期: {exc}. JSON 片段: {json_text[:500]}"
            ) from exc


def _extract_json_object(text: str) -> str:
    fenced = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, re.S)
    if fenced:
        return fenced.group(1)

    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        return text[start : end + 1]

    return text
