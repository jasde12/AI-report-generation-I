from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from app.config import PROJECT_ROOT


@dataclass(frozen=True)
class TemplateDefinition:
    template_id: str
    name: str
    description: str
    file_glob: str | None = None


_TEMPLATES: dict[str, TemplateDefinition] = {
    "general_v1": TemplateDefinition(
        template_id="general_v1",
        name="重要急救責任醫院診斷輔導報告模板",
        description=(
            "以 D診斷輔導報告(OOOO醫院)v01_範本.docx 的封面、章節順序、"
            "段落風格與附件編排為基礎生成正式報告。"
        ),
        file_glob="D診斷輔導報告*範本.docx",
    ),
}


def get_template_definition(template_id: str) -> TemplateDefinition:
    template = _TEMPLATES.get(template_id)
    if template is None:
        supported = ", ".join(sorted(_TEMPLATES))
        raise ValueError(f"不支援的 template_id: {template_id}。目前支援: {supported}")
    return template


def get_template_docx_path(template_id: str) -> Path | None:
    template = get_template_definition(template_id)
    if not template.file_glob:
        return None

    matches = sorted(PROJECT_ROOT.glob(template.file_glob))
    if not matches:
        raise FileNotFoundError(
            f"找不到 template_id={template_id} 對應的 docx 範本檔，搜尋條件: {template.file_glob}"
        )

    return matches[0]


def list_template_ids() -> list[str]:
    return sorted(_TEMPLATES)
