from __future__ import annotations

import json
from typing import Any

from app.models.schemas import ReportJSON


REPORT_SCHEMA_NAME = "report_json"


def get_report_json_schema() -> dict[str, Any]:
    return ReportJSON.model_json_schema()


def get_report_schema_outline() -> str:
    return json.dumps(get_report_json_schema(), ensure_ascii=False, indent=2)
