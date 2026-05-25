from __future__ import annotations

from typing import Any

from pydantic import BaseModel, ConfigDict, Field


class QuestionnaireBaseModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class FieldValue(QuestionnaireBaseModel):
    field_id: str
    label: str
    value: Any | None = None
    unit: str | None = None
    filled: bool = False
    raw_text: str | None = None


class CheckboxOption(QuestionnaireBaseModel):
    label: str
    selected: bool
    value: Any | None = None
    unit: str | None = None


class CheckboxField(QuestionnaireBaseModel):
    question: str
    selection_type: str = "multiple"
    options: list[CheckboxOption] = Field(default_factory=list)
    selected_values: list[str] = Field(default_factory=list)
    raw_text: str | None = None


class CalculationCheck(QuestionnaireBaseModel):
    field: str
    document_value: float | str | None = None
    system_calculated_value: float | str | None = None
    consistent: bool
    note: str = ""


class DataQuality(QuestionnaireBaseModel):
    missing_fields: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    calculation_checks: list[CalculationCheck] = Field(default_factory=list)


class QuestionnaireParseResponse(QuestionnaireBaseModel):
    raw_json: dict[str, Any]
    normalized_json: dict[str, Any]
    report_input_json: dict[str, Any]
