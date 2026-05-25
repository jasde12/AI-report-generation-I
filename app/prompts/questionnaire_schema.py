from __future__ import annotations

import json
from typing import Any

from app.models.schemas import HospitalResilienceDiagnosisForm


QUESTIONNAIRE_SCHEMA_NAME = "hospital_resilience_diagnosis_form"


def get_questionnaire_understanding_schema() -> dict[str, Any]:
    return HospitalResilienceDiagnosisForm.model_json_schema()


def get_questionnaire_schema_outline() -> str:
    return json.dumps(get_questionnaire_understanding_schema(), ensure_ascii=False, indent=2)
