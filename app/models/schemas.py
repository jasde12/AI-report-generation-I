from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator


class StrictBaseModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class CheckboxValue(StrictBaseModel):
    field_name: str
    value: str
    confidence_note: str | None = None


class TableCell(StrictBaseModel):
    row_index: int
    column_index: int
    text: str
    row_span: int = 1
    column_span: int = 1


class MergedCellRange(StrictBaseModel):
    start_row: int
    start_column: int
    end_row: int
    end_column: int


class QuestionOption(StrictBaseModel):
    label: str
    selected: bool


class ExtractedQuestion(StrictBaseModel):
    question_id: str
    question: str
    selection_type: Literal["multiple"] = "multiple"
    options: list[QuestionOption] = Field(default_factory=list)
    selected_values: list[str] = Field(default_factory=list)


class SourceBlock(StrictBaseModel):
    order: int
    block_type: Literal["paragraph", "table"]
    text: str | None = None
    table_index: int | None = None
    table_title: str | None = None
    table_markdown: str | None = None


class ExtractedTable(StrictBaseModel):
    title: str | None = None
    sheet_name: str | None = None
    page_number: int | None = None
    table_kind: str = "generic_matrix"
    columns: list[str] = Field(default_factory=list)
    rows: list[Any] = Field(default_factory=list)
    raw_rows: list[list[Any]] = Field(default_factory=list)
    header_rows: list[list[str]] = Field(default_factory=list)
    semantic_lines: list[str] = Field(default_factory=list)
    questions: list[ExtractedQuestion] = Field(default_factory=list)
    markdown: str = ""
    row_count: int = 0
    column_count: int = 0
    cells: list[TableCell] = Field(default_factory=list)
    merged_ranges: list[MergedCellRange] = Field(default_factory=list)
    header_depth: int = 0
    warnings: list[str] = Field(default_factory=list)


class ImageAnalysisResult(StrictBaseModel):
    summary: str
    detected_items: list[str] = Field(default_factory=list)
    checkbox_values: list[CheckboxValue] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    uncertain: bool = False


class ExtractedSource(StrictBaseModel):
    type: Literal["docx", "table", "pdf", "image"]
    source_name: str
    paragraphs: list[str] = Field(default_factory=list)
    text: str | None = None
    tables: list[ExtractedTable] = Field(default_factory=list)
    blocks: list[SourceBlock] = Field(default_factory=list)
    mime_type: str | None = None
    data_url: str | None = None
    warnings: list[str] = Field(default_factory=list)
    metadata: dict[str, Any] = Field(default_factory=dict)
    image_analysis: ImageAnalysisResult | None = None


class NormalizedFact(StrictBaseModel):
    source_name: str
    fact_type: str
    content: str
    block_order: int | None = None


class NormalizedDocument(StrictBaseModel):
    project_name: str
    hospital_name: str | None = None
    template_id: str
    sources: list[ExtractedSource] = Field(default_factory=list)
    facts: list[NormalizedFact] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class EvidenceReference(StrictBaseModel):
    source_name: str
    fact_type: str
    content: str
    block_order: int | None = None

    @model_validator(mode="before")
    @classmethod
    def coerce_legacy_string(cls, value: Any) -> Any:
        if isinstance(value, str):
            return {
                "source_name": "",
                "fact_type": "legacy_text",
                "content": value,
                "block_order": None,
            }
        return value


class EvidenceText(StrictBaseModel):
    text: str
    evidence: list[EvidenceReference] = Field(default_factory=list)

    @model_validator(mode="before")
    @classmethod
    def coerce_legacy_string(cls, value: Any) -> Any:
        if isinstance(value, str):
            return {"text": value, "evidence": []}
        return value


class KeyFinding(StrictBaseModel):
    item: str
    description: str
    evidence: list[EvidenceReference] = Field(default_factory=list)


class ReportTable(StrictBaseModel):
    title: str
    columns: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)


class ReportSection(StrictBaseModel):
    title: str
    paragraphs: list[EvidenceText] = Field(default_factory=list)


class ReportJSON(StrictBaseModel):
    title: str
    summary: str
    background: str
    key_findings: list[KeyFinding] = Field(default_factory=list)
    tables: list[ReportTable] = Field(default_factory=list)
    recommendations: list[EvidenceText] = Field(default_factory=list)
    missing_information: list[str] = Field(default_factory=list)
    introduction_paragraphs: list[str] = Field(default_factory=list)
    diagnosis_paragraphs: list[EvidenceText] = Field(default_factory=list)
    recommendation_intro: str = ""
    recommendation_sections: list[ReportSection] = Field(default_factory=list)


class DiagnosisSourceFile(StrictBaseModel):
    file_name: str | None = None
    file_type: str | None = None


class DiagnosisBasicInfo(StrictBaseModel):
    organizer: str | None = None
    commissioned_unit: str | None = None
    hospital_name: str | None = None
    diagnosis_date: str | None = None
    diagnosis_expert: str | None = None


class DiagnosisCalculation(StrictBaseModel):
    calculation_name: str = ""
    scenario: str | None = None
    formula: str | None = None
    inputs: dict[str, Any] = Field(default_factory=dict)
    document_result_hours: float | None = None
    document_result_days: float | None = None
    system_calculated_result_hours: float | None = None
    system_calculated_result_days: float | None = None
    target_hours: float | None = None
    target_days: float | None = None
    meets_target: bool | None = None
    meets_72_hours: bool | None = None
    note: str | None = None


class DiagnosisSection(StrictBaseModel):
    section_id: str = ""
    section_name: str = ""
    current_status: dict[str, Any] = Field(default_factory=dict)
    extracted_facts: dict[str, Any] = Field(default_factory=dict)
    calculations: list[DiagnosisCalculation] = Field(default_factory=list)
    risk_points: list[str] = Field(default_factory=list)
    recommendations: list[str] = Field(default_factory=list)


class OverallSummaryItem(StrictBaseModel):
    can_meet_72_hours: bool | None = None
    summary: str = ""


class OverallSummary(StrictBaseModel):
    power: OverallSummaryItem | None = None
    water: OverallSummaryItem | None = None
    medical_gas: OverallSummaryItem | None = None
    it_backup: OverallSummaryItem | None = None
    facility_maintenance: OverallSummaryItem | None = None
    central_monitoring: OverallSummaryItem | None = None


class PossibleInconsistency(StrictBaseModel):
    field: str = ""
    document_value: str | None = None
    system_calculated_value: str | None = None
    system_interpretation: str | None = None
    note: str = ""


class DataQuality(StrictBaseModel):
    missing_or_blank_fields: list[str] = Field(default_factory=list)
    possible_inconsistencies: list[PossibleInconsistency] = Field(default_factory=list)
    placeholder_fields: list[str] = Field(default_factory=list)


class HospitalResilienceDiagnosisForm(StrictBaseModel):
    document_type: Literal["hospital_resilience_diagnosis_form"] = (
        "hospital_resilience_diagnosis_form"
    )
    project_name: str = ""
    form_name: str = "急救責任醫院孤島效應下能資源供應韌性問卷診斷建議單"
    target_hours: int = 72
    source_file: DiagnosisSourceFile = Field(default_factory=DiagnosisSourceFile)
    basic_info: DiagnosisBasicInfo = Field(default_factory=DiagnosisBasicInfo)
    sections: list[DiagnosisSection] = Field(default_factory=list)
    overall_summary: OverallSummary = Field(default_factory=OverallSummary)
    data_quality: DataQuality = Field(default_factory=DataQuality)


class GenerateReportResponse(StrictBaseModel):
    success: bool
    normalized_preview: dict[str, Any]
    report_json: ReportJSON
    output_file: str
    download_url: str


class ValidateQuestionnaireResponse(StrictBaseModel):
    success: bool
    normalized_preview: dict[str, Any]
    understanding_json: dict[str, Any]
