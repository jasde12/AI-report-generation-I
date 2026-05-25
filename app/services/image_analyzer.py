from __future__ import annotations

from functools import cached_property

from openai import OpenAI

from app.config import get_settings
from app.models.schemas import ExtractedSource, ImageAnalysisResult
from app.prompts.vision_prompts import VISION_SYSTEM_PROMPT, build_vision_user_prompt


class ImageAnalyzer:
    def __init__(self, model: str | None = None) -> None:
        self.settings = get_settings()
        self.model = model or self.settings.openai_model

    @cached_property
    def client(self) -> OpenAI:
        if not self.settings.openai_api_key:
            raise RuntimeError("OPENAI_API_KEY 尚未設定。")
        return OpenAI(api_key=self.settings.openai_api_key)

    def analyze(self, source: ExtractedSource) -> ImageAnalysisResult:
        if source.type != "image" or not source.data_url:
            raise ValueError("ImageAnalyzer 只接受帶有 data_url 的圖片來源。")

        try:
            response = self.client.responses.parse(
                model=self.model,
                instructions=VISION_SYSTEM_PROMPT,
                input=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "input_text",
                                "text": build_vision_user_prompt(source.source_name),
                            },
                            {
                                "type": "input_image",
                                "image_url": source.data_url,
                                "detail": "high",
                            },
                        ],
                    }
                ],
                text_format=ImageAnalysisResult,
                temperature=0,
                max_output_tokens=1200,
            )
        except Exception as exc:
            raise RuntimeError(f"OpenAI 圖片分析呼叫失敗: {exc}") from exc

        parsed = response.output_parsed
        if parsed is None:
            raise RuntimeError("OpenAI 沒有回傳可解析的圖片分析結果。")

        return parsed
