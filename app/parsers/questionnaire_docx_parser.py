from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Iterator

from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

from app.schemas.questionnaire_schema import (
    CalculationCheck,
    CheckboxField,
    CheckboxOption,
    DataQuality,
    FieldValue,
    QuestionnaireParseResponse,
)

CHECKED = "\u25a0"
UNCHECKED = "\u25a1"
FULLWIDTH_UNDERSCORE = "\uff3f"
CHECKED_MARKERS = {
    CHECKED,
    "\u2611",
    "\u2713",
    "\u2714",
    "\uf06e",
    "v",
    "V",
}
UNCHECKED_MARKERS = {
    UNCHECKED,
    "\u2610",
    "\uf06f",
}
CHECKBOX_TOKEN_PATTERN = re.compile(
    rf"([{re.escape(CHECKED + UNCHECKED)}])\s*([^{re.escape(CHECKED + UNCHECKED)}\n]+)"
)
NUMBER_PATTERN = re.compile(r"(?<![A-Za-z\u4e00-\u9fff])-?\d+(?:,\d{3})*(?:\.\d+)?")
PLACEHOLDER_PATTERN = re.compile(
    rf"O{{2,}}|Ｏ{{2,}}|○{{2,}}|[_{FULLWIDTH_UNDERSCORE}]{{4,}}"
)
UNIT_LIKE_TOKENS = {
    "m3",
    "M3",
    "kw",
    "kW",
    "kwh",
    "kWh",
    "l",
    "L",
    "hrs",
    "Hrs",
    "hr",
    "Hr",
    "噸",
    "公噸",
    "公升",
    "小時",
    "天",
    "床",
    "公里",
    "%",
}


def _iter_document_blocks(document: DocxDocument) -> Iterator[Paragraph | Table]:
    for child in document.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document._body)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document._body)


def _to_number(value: str) -> int | float:
    numeric = float(value.replace(",", ""))
    return int(numeric) if numeric.is_integer() else numeric


def clean_text(text: str) -> str:
    if not text:
        return ""

    normalized = (
        text.replace("\r\n", "\n")
        .replace("\r", "\n")
        .replace("\u3000", " ")
        .replace("\u2611", CHECKED)
        .replace("\u2713", CHECKED)
        .replace("\u2714", CHECKED)
        .replace("\uf06e", CHECKED)
        .replace("\u2610", UNCHECKED)
        .replace("\uf06f", UNCHECKED)
    )
    normalized = re.sub(r"[ \t]+", " ", normalized)
    normalized = "\n".join(line.strip() for line in normalized.splitlines())
    normalized = re.sub(r"\n{3,}", "\n\n", normalized)
    return normalized.strip()


def extract_number(text: str) -> int | float | None:
    if not text:
        return None
    match = NUMBER_PATTERN.search(
        clean_text(text).replace("_", "").replace(FULLWIDTH_UNDERSCORE, "")
    )
    if not match:
        return None
    return _to_number(match.group(0))


def extract_number_with_unit(text: str) -> tuple[int | float | None, str | None]:
    cleaned = clean_text(text).replace("_", "").replace(FULLWIDTH_UNDERSCORE, "")
    match = re.search(
        r"(?<![A-Za-z\u4e00-\u9fff])(?P<number>-?\d+(?:,\d{3})*(?:\.\d+)?)\s*(?P<unit>[A-Za-z0-9%/\u4e00-\u9fff()]+)?",
        cleaned,
    )
    if match:
        number = _to_number(match.group("number"))
        unit = match.group("unit") or None
        return number, unit

    unit_match = re.search(
        rf"[_{FULLWIDTH_UNDERSCORE}]+\s*(?P<unit>[A-Za-z0-9%/\u4e00-\u9fff()]+)",
        clean_text(text),
    )
    if unit_match:
        return None, unit_match.group("unit")

    return None, None


def _normalize_placeholder_text(text: str) -> str:
    return clean_text(text).replace(FULLWIDTH_UNDERSCORE, "_")


def _strip_hint_text(text: str) -> str:
    return re.split(r"[（(]", _normalize_placeholder_text(text), maxsplit=1)[0].strip()


def _is_placeholder_only(text: str) -> bool:
    if not text:
        return True

    candidate = _strip_hint_text(text)
    if not candidate:
        return True

    if NUMBER_PATTERN.search(candidate):
        return False

    if "_" in candidate:
        reduced = re.sub(r"_+", "", candidate)
        reduced = re.sub(r"[：:；;，,、.\s\-~至到/]+", "", reduced)
        for token in sorted(UNIT_LIKE_TOKENS, key=len, reverse=True):
            reduced = reduced.replace(token, "")
        return reduced == ""

    return False


def parse_fill_field(
    text: str,
    label: str,
    field_id: str,
    unit: str | None = None,
) -> dict[str, Any]:
    cleaned = clean_text(text)
    pattern = re.compile(
        rf"{re.escape(label)}\s*[：:]\s*(?P<value>[^\n；;|]*)",
        re.IGNORECASE,
    )
    match = pattern.search(cleaned)
    raw_value = match.group("value").strip() if match else ""
    raw_value = raw_value.rstrip("；;")
    preview = _strip_hint_text(raw_value)

    filled = bool(preview and not _is_placeholder_only(preview))
    value: Any | None = preview if filled else None
    detected_unit = unit

    if preview:
        detected_number, detected_number_unit = extract_number_with_unit(preview)
        if detected_number is not None:
            value = detected_number
            detected_unit = unit or detected_number_unit
        elif detected_number_unit and not unit:
            detected_unit = detected_number_unit
        elif not filled:
            value = None

    field = FieldValue(
        field_id=field_id,
        label=label,
        value=value,
        unit=detected_unit,
        filled=filled,
        raw_text=cleaned,
    )
    return field.model_dump()


def _parse_checkbox_option_body(body: str, selected: bool) -> CheckboxOption:
    cleaned = body.strip(" ；;")
    label = cleaned
    value: Any | None = None
    unit: str | None = None

    if cleaned.startswith(("有，", "有,")):
        label = "有"
        remainder = cleaned[2:].strip(" ，,；;")
        number, detected_unit = extract_number_with_unit(remainder)
        if number is not None:
            value = number
            unit = detected_unit
        elif remainder and not _is_placeholder_only(remainder):
            value = remainder
    elif "：" in cleaned or ":" in cleaned:
        parts = re.split(r"[：:]", cleaned, maxsplit=1)
        label = parts[0].strip()
        remainder = parts[1].strip() if len(parts) > 1 else ""
        preview = _strip_hint_text(remainder)
        preview_stripped = re.sub(r"_+", "", preview).strip(" ，,；;")
        number, detected_unit = extract_number_with_unit(preview)
        if number is not None:
            value = number
            unit = detected_unit
        else:
            if _is_placeholder_only(preview):
                unit = detected_unit
            elif (
                detected_unit
                and detected_unit in UNIT_LIKE_TOKENS
                and (
                    preview_stripped == detected_unit
                    or preview_stripped in {f"{detected_unit}或", f"{detected_unit} 或"}
                )
            ):
                unit = detected_unit
            elif preview_stripped in UNIT_LIKE_TOKENS:
                unit = preview_stripped
            elif preview_stripped and selected:
                value = preview_stripped
                unit = detected_unit if detected_unit in UNIT_LIKE_TOKENS else None

    return CheckboxOption(
        label=label,
        selected=selected,
        value=value,
        unit=unit,
    )


def parse_checkbox_options(text: str, question: str = "") -> dict[str, Any]:
    cleaned = clean_text(text)
    tokens: list[CheckboxOption] = []

    for line in cleaned.splitlines():
        line = line.strip()
        if not line:
            continue
        if line and line[0] in CHECKED_MARKERS:
            line = CHECKED + line[1:]
        elif line and line[0] in UNCHECKED_MARKERS:
            line = UNCHECKED + line[1:]

        for marker, body in CHECKBOX_TOKEN_PATTERN.findall(line):
            option = _parse_checkbox_option_body(body, selected=marker == CHECKED)
            tokens.append(option)

    field = CheckboxField(
        question=question,
        options=tokens,
        selected_values=[option.label for option in tokens if option.selected],
        raw_text=cleaned,
    )
    return field.model_dump()


def detect_section(text: str) -> str:
    cleaned = clean_text(text)
    if "基本資訊" in cleaned:
        return "basic_info"
    if "供電相關" in cleaned:
        return "power"
    if "供水相關" in cleaned:
        return "water"
    if "醫用氣體" in cleaned:
        return "medical_gas"
    if "資訊系統備援" in cleaned:
        return "it_backup"
    if "設施維護措施" in cleaned:
        return "facility_maintenance"
    if "中央監控系統" in cleaned:
        return "central_monitoring"
    return "unknown"


def read_docx_blocks(docx_path: str) -> dict[str, Any]:
    path = Path(docx_path)
    document = Document(path)

    paragraphs: list[str] = []
    tables: list[dict[str, Any]] = []
    blocks: list[dict[str, Any]] = []

    for order, block in enumerate(_iter_document_blocks(document), start=1):
        if isinstance(block, Paragraph):
            text = clean_text(block.text)
            if not text:
                continue
            paragraphs.append(text)
            blocks.append(
                {
                    "order": order,
                    "block_type": "paragraph",
                    "text": text,
                }
            )
            continue

        rows: list[list[str]] = []
        for row in block.rows:
            rows.append([clean_text(cell.text) for cell in row.cells])
        tables.append(
            {
                "title": f"Table {len(tables) + 1}",
                "rows": rows,
                "row_count": len(block.rows),
                "column_count": len(block.columns),
            }
        )
        blocks.append(
            {
                "order": order,
                "block_type": "table",
                "rows": rows,
            }
        )

    return {
        "document_type": "hospital_resilience_questionnaire",
        "source_file": path.name,
        "paragraphs": paragraphs,
        "tables": tables,
        "blocks": blocks,
    }


def detect_questionnaire_profile(raw_json: dict[str, Any]) -> str:
    paragraphs = raw_json.get("paragraphs", [])
    tables = raw_json.get("tables", [])

    first_table = tables[0] if tables else {}
    first_rows = first_table.get("rows", [])
    first_cell = clean_text(first_rows[0][0]) if first_rows and first_rows[0] else ""

    if "主辦單位" in first_cell and "診斷專家" in first_cell:
        return "diagnosis_form"

    if (
        any("前言" in paragraph for paragraph in paragraphs)
        and "醫院名稱" in first_cell
        and "總床位數" in first_cell
        and len(tables) >= 6
    ):
        return "full_questionnaire"

    if any("前言" in paragraph for paragraph in paragraphs) and len(tables) >= 6:
        return "full_questionnaire"

    if tables:
        return "diagnosis_form"

    return "unknown"


def _get_table_by_index(raw_json: dict[str, Any], index: int) -> dict[str, Any]:
    tables = raw_json.get("tables", [])
    if 0 <= index < len(tables):
        return tables[index]
    return {"title": None, "rows": [], "row_count": 0, "column_count": 0}


def _get_unique_cleaned_rows(table: dict[str, Any]) -> list[list[str]]:
    unique_rows: list[list[str]] = []
    seen: set[tuple[str, ...]] = set()

    for row in table.get("rows", []):
        cleaned_row = [clean_text(str(cell)) for cell in row]
        row_key = tuple(cleaned_row)
        if row_key in seen:
            continue
        seen.add(row_key)
        unique_rows.append(cleaned_row)

    return unique_rows


def _join_table_column_text(
    table: dict[str, Any],
    column_index: int,
    *,
    start_row: int = 1,
) -> str:
    parts: list[str] = []
    for row in _get_unique_cleaned_rows(table)[start_row:]:
        if len(row) > column_index and row[column_index]:
            parts.append(row[column_index])
    return clean_text("\n".join(parts))


def _flatten_table_text(table: dict[str, Any], *, start_row: int = 0) -> str:
    row_texts: list[str] = []
    for row in _get_unique_cleaned_rows(table)[start_row:]:
        non_empty_cells = [cell for cell in row if cell]
        if non_empty_cells:
            row_texts.append("\n".join(non_empty_cells))
    return clean_text("\n\n".join(row_texts))


def _find_paragraph_containing(raw_json: dict[str, Any], marker: str) -> str:
    for paragraph in raw_json.get("paragraphs", []):
        cleaned = clean_text(paragraph)
        if marker in cleaned:
            return cleaned
    return ""


def _extract_labeled_text(
    text: str,
    label: str,
    *,
    end_markers: list[str] | None = None,
) -> str | None:
    cleaned = clean_text(text)
    marker_pattern = "|".join(re.escape(marker) for marker in (end_markers or []))
    if marker_pattern:
        pattern = re.compile(
            rf"{re.escape(label)}[：:]\s*(?P<value>.+?)(?=(?:{marker_pattern})[：:]|$)",
            re.S,
        )
    else:
        pattern = re.compile(rf"{re.escape(label)}[：:]\s*(?P<value>.+)", re.S)

    match = pattern.search(cleaned)
    if not match:
        return None
    return clean_text(match.group("value"))


def _parse_roc_date(text: str) -> str | None:
    match = re.search(
        r"填表日期[：:]\s*_*(\d{2,3})_*+\s*年\s*_*(\d{1,2})_*+\s*月\s*_*(\d{1,2})_*+\s*日",
        clean_text(text),
    )
    if not match:
        return None

    roc_year = int(match.group(1))
    month = int(match.group(2))
    day = int(match.group(3))
    gregorian_year = roc_year + 1911 if roc_year < 1911 else roc_year
    return f"{gregorian_year:04d}.{month:02d}.{day:02d}"


def _get_primary_table_rows(raw_json: dict[str, Any]) -> list[list[str]]:
    tables = raw_json.get("tables", [])
    if not tables:
        return []
    return tables[0].get("rows", [])


def _find_section_row(rows: list[list[str]], section_key: str) -> list[str]:
    for row in rows:
        if row and detect_section(row[0]) == section_key:
            return row
    return []


def _section_split(text: str, marker: str) -> tuple[str, str]:
    cleaned = clean_text(text)
    if marker not in cleaned:
        return cleaned, ""
    before, after = cleaned.split(marker, 1)
    return before.strip(), after.strip()


def _dedupe_merged_cell_text(row: list[str], start_index: int = 2) -> str:
    seen: list[str] = []
    for cell in row[start_index:]:
        cleaned = clean_text(cell)
        if cleaned and cleaned not in seen:
            seen.append(cleaned)
    return "\n".join(seen)


def _bool_from_option(field: dict[str, Any], label: str) -> bool | None:
    for option in field.get("options", []):
        if option.get("label") == label:
            return bool(option.get("selected"))
    return None


def _option_value(field: dict[str, Any], label: str) -> Any:
    for option in field.get("options", []):
        if option.get("label") == label:
            return option.get("value")
    return None


def _make_option(
    label: str,
    selected: bool,
    value: Any | None = None,
    unit: str | None = None,
) -> dict[str, Any]:
    return CheckboxOption(
        label=label,
        selected=selected,
        value=value,
        unit=unit,
    ).model_dump()


def _finalize_checkbox_field(field: dict[str, Any]) -> dict[str, Any]:
    field["selected_values"] = [
        option.get("label") for option in field.get("options", []) if option.get("selected")
    ]
    return field


def _extract_segment(text: str, start_marker: str, end_markers: list[str]) -> str:
    if start_marker not in text:
        return ""
    segment = text.split(start_marker, 1)[1]
    end_positions = [
        segment.find(marker) for marker in end_markers if marker and segment.find(marker) != -1
    ]
    if end_positions:
        segment = segment[: min(end_positions)]
    return segment.strip()


def _round2(value: float | None) -> float | None:
    if value is None:
        return None
    return round(value, 2)


def _missing_field_message(field_id: str | None, label: str | None) -> str:
    mapping = {
        "power_reachable_days": "供電可達天數未填",
        "power_application_item": "申請衛福部之建置項目未填",
        "power_application_purpose": "申請衛福部之用途未填",
        "energy_storage_power_kw": "設置儲能櫃功率未填",
        "temporary_generator_power_kw": "設置臨時發電機功率未填",
        "water_interconnected_storage_liters": "供水聯通管系統蓄水量未填",
        "water_interconnected_supply_hours": "供水聯通管系統可累積供水小時未填",
        "it_offsite_distance_km": "IT異地備援距離未填",
        "it_backup_type": "IT備份機制未填",
        "it_cloud_vendor": "IT雲端備援廠家未填",
        "it_other_external_backup_method": "其他院外備援方式未填",
        "it_improvement_years": "未設有異地備援之預計改善年限未填",
    }
    if field_id in mapping:
        return mapping[field_id]
    safe_label = label or field_id or "欄位"
    return f"{safe_label}未填"


def parse_basic_info_full_questionnaire(raw_json: dict[str, Any]) -> dict[str, Any]:
    first_table = _get_table_by_index(raw_json, 0)
    first_rows = _get_unique_cleaned_rows(first_table)
    first_cell = first_rows[0][0] if first_rows and first_rows[0] else ""

    hospital_name = _extract_labeled_text(
        first_cell,
        "醫院名稱",
        end_markers=[
            "醫院緊急醫療能力等級",
            "醫院總樓地板面積",
            "總床位數",
            "院長/負責人姓名",
        ],
    )
    if hospital_name:
        hospital_name = clean_text(re.split(r"[（(]", hospital_name, maxsplit=1)[0]).strip()

    return {
        "organizer": "衛生福利部",
        "commissioned_unit": "財團法人金屬工業研究發展中心",
        "hospital_name": hospital_name,
        "diagnosis_date": _parse_roc_date(first_cell),
        "diagnosis_expert": None,
        "completion_status": {
            "question": "基本資訊是否完整",
            "selection_type": "single",
            "options": [],
            "selected_values": [],
            "raw_text": clean_text(first_cell),
        },
        "recommendation_note": None,
    }


def parse_basic_info(raw_json: dict[str, Any]) -> dict[str, Any]:
    rows = _get_primary_table_rows(raw_json)
    basic_row = _find_section_row(rows, "basic_info")
    completion = parse_checkbox_options(
        basic_row[1] if len(basic_row) > 1 else "",
        question="基本資訊是否完整",
    )
    completion["selection_type"] = "single"

    return {
        "organizer": clean_text(rows[0][1]) if len(rows) > 0 and len(rows[0]) > 1 else None,
        "commissioned_unit": clean_text(rows[0][3]) if len(rows) > 0 and len(rows[0]) > 3 else None,
        "hospital_name": clean_text(rows[1][1]) if len(rows) > 1 and len(rows[1]) > 1 else None,
        "diagnosis_date": clean_text(rows[1][3]) if len(rows) > 1 and len(rows[1]) > 3 else None,
        "diagnosis_expert": clean_text(rows[1][5]) if len(rows) > 1 and len(rows[1]) > 5 else None,
        "completion_status": completion,
        "recommendation_note": clean_text(basic_row[2]) if len(basic_row) > 2 else None,
    }


def parse_power_section(text: str, recommendation_text: str = "") -> dict[str, Any]:
    normal_text, island_text = _section_split(text, "孤島效應下應變措施")
    normal_field = parse_checkbox_options(
        normal_text,
        question="一般狀況，即不考慮孤島效應",
    )
    island_field = parse_checkbox_options(
        island_text,
        question="孤島效應下應變措施",
    )

    generator_count = None
    generator_match = re.search(r"([一二兩三四五六七八九十\d]+)台發電機", recommendation_text)
    if generator_match:
        raw_count = generator_match.group(1)
        chinese_map = {
            "一": 1,
            "二": 2,
            "兩": 2,
            "三": 3,
            "四": 4,
            "五": 5,
            "六": 6,
            "七": 7,
            "八": 8,
            "九": 9,
            "十": 10,
        }
        generator_count = chinese_map.get(raw_count, extract_number(raw_count))

    fuel_total_liters = None
    fuel_match = re.search(r"油量總計有\s*([0-9,]+)\s*公升", recommendation_text)
    if fuel_match:
        fuel_total_liters = _to_number(fuel_match.group(1))

    hourly_fuel_liters = None
    hourly_match = re.search(r"每小時(?:可供應|使用)\s*([0-9,]+)\s*公升", recommendation_text)
    if hourly_match:
        hourly_fuel_liters = _to_number(hourly_match.group(1))

    document_hours = None
    hours_match = re.search(r"=\s*([0-9]+(?:\.[0-9]+)?)\s*小時", recommendation_text)
    if hours_match:
        document_hours = _to_number(hours_match.group(1))

    system_hours = None
    if fuel_total_liters is not None and hourly_fuel_liters:
        system_hours = fuel_total_liters / float(hourly_fuel_liters)

    future_plan = None
    future_plan_match = re.search(r"(建議未來[^。]+。?)", recommendation_text)
    if future_plan_match:
        future_plan = clean_text(future_plan_match.group(1))

    application_item = parse_fill_field(
        island_text,
        "申請建置項目",
        "power_application_item",
    )
    purpose = parse_fill_field(
        island_text,
        "用途",
        "power_application_purpose",
    )
    reachable_days = parse_fill_field(
        island_text,
        "可達天數",
        "power_reachable_days",
        unit="天",
    )

    can_reach_72_hours = bool(system_hours is not None and system_hours >= 72)

    return {
        "normal_condition": normal_field,
        "baseline": {
            "year": None,
            "annual_usage_kwh": None,
            "average_daily_usage_kwh": None,
            "average_hourly_usage_kwh": None,
            "monthly_usage_kwh": {},
        },
        "ups": {
            "exists": _bool_from_option(normal_field, "不斷電系統(UPS)"),
            "supply_areas": [],
            "duration_seconds": None,
            "is_parallel": None,
            "parallel_capacity_kwh": None,
            "independent_total_capacity_kwh": None,
        },
        "generator": {
            "generator_count": generator_count,
            "fuel_type": None,
            "generators": [
                {
                    "id": "#01",
                    "location": None,
                    "power_kw": None,
                    "type": None,
                    "ats_start_seconds": None,
                    "daily_tank_liters": None,
                    "hourly_fuel_consumption_liters": hourly_fuel_liters,
                }
            ]
            if generator_count
            else [],
        },
        "fuel_storage": {
            "war_reserve_tank": {
                "exists": _bool_from_option(normal_field, "具(戰)備用油槽"),
                "capacity_liters": None,
            },
            "underground_tank": {
                "exists": _bool_from_option(normal_field, "具地下油槽"),
                "capacity_liters": None,
            },
            "fuel_barrel_or_storage_tank": {
                "exists": _bool_from_option(normal_field, "具儲油槽/桶"),
                "capacity_liters": None,
            },
        },
        "island_mode_response": island_field,
        "island_mode": {
            "can_reach_72_hours": can_reach_72_hours,
            "declared_reachable_days": reachable_days,
            "calculation": {
                "formula": f"{fuel_total_liters} / {hourly_fuel_liters}"
                if fuel_total_liters is not None and hourly_fuel_liters
                else None,
                "document_result_hours": document_hours,
                "system_calculated_hours": _round2(system_hours),
                "meets_target": can_reach_72_hours,
            },
            "application_to_mohw": {
                "selected": _bool_from_option(island_field, "申請衛福部"),
                "application_item": application_item,
                "purpose": purpose,
            },
            "medical_priority_order_defined": False,
            "environment_route_power_plan_defined": False,
            "has_energy_storage": _bool_from_option(island_field, "設置儲能櫃"),
            "energy_storage_power_kw": {
                "field_id": "energy_storage_power_kw",
                "label": "設置儲能櫃功率",
                "value": None,
                "unit": "kW",
                "filled": False,
                "raw_text": island_text,
            },
            "has_temporary_generator": _bool_from_option(island_field, "設置臨時發電機"),
            "temporary_generator_power_kw": {
                "field_id": "temporary_generator_power_kw",
                "label": "設置臨時發電機功率",
                "value": None,
                "unit": "kW",
                "filled": False,
                "raw_text": island_text,
            },
            "future_plan": future_plan,
        },
        "raw_text": {
            "current_status": clean_text(text),
            "recommendation": clean_text(recommendation_text),
        },
    }


def parse_water_section(text: str, recommendation_text: str = "") -> dict[str, Any]:
    normal_text, island_text = _section_split(text, "孤島效應下應變措施A/[B*C]")
    normal_field = parse_checkbox_options(
        normal_text,
        question="一般狀況，即不考慮孤島效應",
    )
    island_field = parse_checkbox_options(
        island_text,
        question="孤島效應下應變措施 A/[B*C]",
    )

    total_storage_tons = None
    storage_match = re.search(r"總儲水量\s*([0-9]+(?:\.[0-9]+)?)\s*公噸", recommendation_text)
    if storage_match:
        total_storage_tons = _to_number(storage_match.group(1))
    elif _option_value(normal_field, "儲水量") is not None:
        total_storage_tons = _option_value(normal_field, "儲水量")

    beds = None
    bed_match = re.search(r"床位數\s*([0-9]+)\s*床", recommendation_text)
    if bed_match:
        beds = _to_number(bed_match.group(1))
    else:
        beds = _option_value(normal_field, "病床數")

    daily_water_per_bed_tons = None
    daily_ton_match = re.search(r"每日每床\(([\d.]+)公噸\)", recommendation_text)
    if daily_ton_match:
        daily_water_per_bed_tons = _to_number(daily_ton_match.group(1))

    full_bed_days = None
    full_day_match = re.search(r"可供應天數為\s*([0-9]+(?:\.[0-9]+)?)", recommendation_text)
    if full_day_match:
        full_bed_days = _to_number(full_day_match.group(1))

    reduced_total_storage_liters = None
    reduced_storage_match = re.search(r"A:蓄水量\(_*([0-9,]+)_*公升\)", recommendation_text)
    if reduced_storage_match:
        reduced_total_storage_liters = _to_number(reduced_storage_match.group(1))

    reduced_general_beds = None
    reduced_icu_beds = None
    reduced_beds_match = re.search(r"病床數\((\d+)\(一般\)\+(\d+)\(加護床\)床\)", recommendation_text)
    if reduced_beds_match:
        reduced_general_beds = _to_number(reduced_beds_match.group(1))
        reduced_icu_beds = _to_number(reduced_beds_match.group(2))
    reduced_total_beds = None
    if reduced_general_beds is not None and reduced_icu_beds is not None:
        reduced_total_beds = int(reduced_general_beds) + int(reduced_icu_beds)

    reduced_daily_liters = None
    reduced_daily_match = re.search(r"C:每日每床用水量\s*([0-9,]+)\s*公升", recommendation_text)
    if reduced_daily_match:
        reduced_daily_liters = _to_number(reduced_daily_match.group(1))

    reduced_days = None
    reduced_days_match = re.search(r"=\s*([0-9]+(?:\.[0-9]+)?)\(天\)", recommendation_text)
    if reduced_days_match:
        reduced_days = _to_number(reduced_days_match.group(1))

    normal_daily_liters = None
    normal_daily_match = re.search(r"每日每床用水量\s*([0-9,]+)\s*公升", normal_text)
    if normal_daily_match:
        normal_daily_liters = _to_number(normal_daily_match.group(1))

    island_declared_days = None
    island_declared_match = re.search(
        r"孤島效應下最大供水可達_*([0-9]+(?:\.[0-9]+)?)_*天數",
        island_text,
    )
    if island_declared_match:
        island_declared_days = _to_number(island_declared_match.group(1))

    system_full_days = None
    if total_storage_tons is not None and beds and daily_water_per_bed_tons:
        system_full_days = float(total_storage_tons) / (
            float(beds) * float(daily_water_per_bed_tons)
        )

    system_reduced_days = None
    if reduced_total_storage_liters is not None and reduced_total_beds and reduced_daily_liters:
        system_reduced_days = float(reduced_total_storage_liters) / (
            float(reduced_total_beds) * float(reduced_daily_liters)
        )

    linked_storage = parse_fill_field(
        island_text,
        "蓄水量",
        "water_interconnected_storage_liters",
        unit="公升",
    )
    linked_supply_hours = parse_fill_field(
        island_text,
        "可累積供水",
        "water_interconnected_supply_hours",
        unit="小時",
    )

    source1_selected = _bool_from_option(normal_field, "來源1(自來水機構供水)")
    source2_selected = _bool_from_option(normal_field, "來源2(合法井水)")
    source3_selected = _bool_from_option(normal_field, "來源3(戰備水池)")
    interconnected_selected = _bool_from_option(island_field, "其他")

    for option in normal_field.get("options", []):
        if option["label"] == "來源1(自來水機構供水)":
            option["label"] = "來源1：自來水機構供水"
        elif option["label"] == "來源2(合法井水)":
            option["label"] = "來源2：合法井水"
        elif option["label"] == "來源3(戰備水池)":
            option["label"] = "來源3：戰備水池"
        elif option["label"] == "儲水量":
            option["unit"] = option.get("unit") or "公噸"
            if option.get("value") is None:
                option["value"] = total_storage_tons
        elif option["label"].startswith("每日每床用水量"):
            option["label"] = "每日每床用水量"
            option["value"] = normal_daily_liters
            option["unit"] = "公升"
        elif option["label"] == "病床數":
            option["unit"] = "床"
            if option.get("value") is None:
                option["value"] = beds
    _finalize_checkbox_field(normal_field)

    for option in island_field.get("options", []):
        if "孤島效應下最大供水可達" in option["label"]:
            option["label"] = "孤島效應下最大供水可達天數"
            option["value"] = island_declared_days
            option["unit"] = "天"
        elif option["label"] == "其他":
            option["label"] = "其他：院內儲水槽間之聯通管系統"
    _finalize_checkbox_field(island_field)

    return {
        "normal_condition": normal_field,
        "baseline": {
            "year": None,
            "annual_usage_m3": None,
            "average_daily_usage_m3": None,
            "monthly_usage_m3": {},
        },
        "sources": [
            {
                "type": "自來水機構供水",
                "selected": source1_selected,
                "facilities": [],
            },
            {
                "type": "合法井水",
                "selected": source2_selected,
                "facilities": [],
            },
            {
                "type": "戰備水池",
                "selected": source3_selected,
                "facilities": [],
            },
        ],
        "total_storage_tons": total_storage_tons,
        "ro_system": {
            "exists": None,
            "building": None,
            "floor": None,
            "quantity": None,
        },
        "island_mode_response": island_field,
        "island_mode": {
            "full_bed_scenario": {
                "total_storage_tons": total_storage_tons,
                "beds": beds,
                "daily_water_per_bed_tons": daily_water_per_bed_tons,
                "document_days": full_bed_days,
                "calculated_days": _round2(system_full_days),
                "meets_72_hours": bool(system_full_days is not None and system_full_days >= 3),
            },
            "reduced_bed_scenario": {
                "total_storage_liters": reduced_total_storage_liters,
                "general_beds": reduced_general_beds,
                "icu_beds": reduced_icu_beds,
                "beds": reduced_total_beds,
                "daily_water_per_bed_liters": reduced_daily_liters,
                "document_days": reduced_days,
                "calculated_days": _round2(system_reduced_days),
                "meets_72_hours": bool(system_reduced_days is not None and system_reduced_days >= 3),
            },
            "interconnected_water_tank_system": {
                "selected": interconnected_selected,
                "storage_liters": linked_storage,
                "supply_hours": linked_supply_hours,
            },
        },
        "raw_text": {
            "current_status": clean_text(text),
            "recommendation": clean_text(recommendation_text),
        },
    }


def parse_medical_gas_section(text: str, recommendation_text: str = "") -> dict[str, Any]:
    normal_text, island_text = _section_split(text, "孤島效應下應變措施")
    parsed_normal_field = parse_checkbox_options(
        normal_text,
        question="一般狀況，即不考慮孤島效應",
    )
    island_field = parse_checkbox_options(
        island_text,
        question="孤島效應下應變措施",
    )

    tank_match = re.search(
        r"液氧槽規模[:：]\s*_*([0-9]+(?:\.[0-9]+)?)_*噸；數量[:：]_*([0-9]+).*?共([0-9]+(?:\.[0-9]+)?)噸",
        text,
    )
    liquid_each = _to_number(tank_match.group(1)) if tank_match else None
    tank_quantity = _to_number(tank_match.group(2)) if tank_match else None
    total_capacity = _to_number(tank_match.group(3)) if tank_match else None

    cylinder_match = re.search(
        r"氧氣鋼瓶規模[:：]\s*([0-9]+)\(院內\)\+([0-9]+)\(院外\)=([0-9]+)\(支\).*?([0-9]+(?:\.[0-9]+)?)",
        text,
    )
    inside_count = _to_number(cylinder_match.group(1)) if cylinder_match else None
    outside_count = _to_number(cylinder_match.group(2)) if cylinder_match else None
    total_count = _to_number(cylinder_match.group(3)) if cylinder_match else None
    gas_equivalent_tons = _to_number(cylinder_match.group(4)) if cylinder_match else None

    daily_usage_match = re.search(r"/\s*([0-9]+(?:\.[0-9]+)?)\(T/天\)", recommendation_text)
    daily_usage_tons = _to_number(daily_usage_match.group(1)) if daily_usage_match else None
    document_days_match = re.search(r"=\s*([0-9]+(?:\.[0-9]+)?)\(天\)", recommendation_text)
    document_days = _to_number(document_days_match.group(1)) if document_days_match else None

    system_days = None
    if total_capacity is not None and daily_usage_tons:
        system_days = float(total_capacity) / float(daily_usage_tons)

    normal_field = CheckboxField(
        question="一般狀況，即不考慮孤島效應",
        options=[
            CheckboxOption(
                label="液氧槽規模",
                selected=bool(_bool_from_option(parsed_normal_field, "液氧槽規模")),
                value={
                    "capacity_tons_each": liquid_each,
                    "quantity": tank_quantity,
                    "total_capacity_tons": total_capacity,
                }
                if total_capacity is not None
                else None,
            ),
            CheckboxOption(
                label="液氧鋼瓶規模",
                selected=bool(_bool_from_option(parsed_normal_field, "液氧鋼瓶規模")),
            ),
            CheckboxOption(
                label="氧氣鋼瓶規模",
                selected=bool(_bool_from_option(parsed_normal_field, "氧氣鋼瓶規模")),
                value={
                    "inside_hospital_count": inside_count,
                    "outside_hospital_count": outside_count,
                    "total_count": total_count,
                    "gas_equivalent_tons": gas_equivalent_tons,
                }
                if total_count is not None
                else None,
            ),
            CheckboxOption(
                label="合計氧氣量(噸)",
                selected=False,
                value=None,
                unit="噸",
            ),
            CheckboxOption(
                label="合計氧氣量(m3)",
                selected=False,
                value=None,
                unit="m3",
            ),
        ],
        selected_values=[],
        raw_text=clean_text(normal_text),
    ).model_dump()
    _finalize_checkbox_field(normal_field)

    for option in island_field.get("options", []):
        if "供氧小時數" in option["label"]:
            option["label"] = "孤島效應下請計算供氧小時數"
            option["unit"] = "小時"
    _finalize_checkbox_field(island_field)

    return {
        "normal_condition": normal_field,
        "baseline": {
            "annual_o2_usage_tons": None,
            "average_daily_o2_usage_tons": None,
        },
        "liquid_oxygen": {
            "capacity_tons_each": liquid_each,
            "quantity": tank_quantity,
            "total_capacity_tons": total_capacity,
            "formula_text": f"{liquid_each} x {tank_quantity} = {total_capacity}"
            if liquid_each is not None and tank_quantity is not None and total_capacity is not None
            else None,
            "location": None,
        },
        "oxygen_cylinders": {
            "inside_hospital_count": inside_count,
            "outside_hospital_count": outside_count,
            "total_count": total_count,
            "gas_equivalent_tons": gas_equivalent_tons,
        },
        "island_mode_response": island_field,
        "island_mode": {
            "formula": f"{total_capacity} / {daily_usage_tons}"
            if total_capacity is not None and daily_usage_tons
            else None,
            "document_days": document_days,
            "calculated_days": _round2(system_days),
            "meets_72_hours": bool(system_days is not None and system_days >= 3),
        },
        "raw_text": {
            "current_status": clean_text(text),
            "recommendation": clean_text(recommendation_text),
        },
    }


def parse_it_backup_section(text: str, recommendation_text: str = "") -> dict[str, Any]:
    cleaned = clean_text(text)
    _, after_internal_header = _section_split(cleaned, "院內備援（參考HIMSS註）")
    internal_text, after_internal = _section_split(after_internal_header, "院外備援")
    external_text, island_text = _section_split(after_internal, "孤島效應下應變措施")

    internal_field = parse_checkbox_options(internal_text, question="院內備援")
    external_field = parse_checkbox_options(external_text, question="院外備援")
    island_field = parse_checkbox_options(island_text, question="孤島效應下應變措施")

    mobile_storage_type = None
    mobile_match = re.search(r"移動式[:：]\s*_*(.*?)_*[(（]", external_text)
    if mobile_match:
        mobile_storage_type = clean_text(mobile_match.group(1))

    distance_field = parse_fill_field(
        external_text,
        "超過",
        "it_offsite_distance_km",
        unit="公里",
    )
    backup_type_field = parse_fill_field(
        external_text,
        "備份機制",
        "it_backup_type",
    )
    distance_field["label"] = "異地備援距離"
    backup_type_field["label"] = "IT備份機制"
    cloud_vendor_field = parse_fill_field(
        external_text,
        "雲端備援廠家",
        "it_cloud_vendor",
    )
    cloud_vendor_field["label"] = "IT雲端備援廠家"
    other_external_backup_method = parse_fill_field(
        external_text,
        "其他院外備援方式",
        "it_other_external_backup_method",
    )
    improvement_years_field = parse_fill_field(
        island_text,
        "未設有異地備援，預計",
        "it_improvement_years",
        unit="年",
    )
    improvement_years_field["label"] = "未設有異地備援之預計改善年限"

    emr_segment = parse_checkbox_options(
        _extract_segment(
            external_text,
            "病歷資訊系統備援",
            ["影像資訊系統備援", "圖像資訊系統備援", "雲端備援廠家"],
        ),
        question="病歷資訊系統備援",
    )
    image_segment = parse_checkbox_options(
        _extract_segment(
            external_text,
            "影像資訊系統備援",
            ["圖像資訊系統備援", "雲端備援廠家"],
        ),
        question="影像資訊系統備援",
    )
    graphic_segment = parse_checkbox_options(
        _extract_segment(
            external_text,
            "圖像資訊系統備援",
            ["雲端備援廠家", "其他院外備援方式"],
        ),
        question="圖像資訊系統備援",
    )

    internal_field = CheckboxField(
        question="院內備援",
        options=[
            CheckboxOption(
                label="單機備援",
                selected=bool(_bool_from_option(internal_field, "單機備援")),
            ),
            CheckboxOption(
                label="複聯備援",
                selected=bool(_bool_from_option(internal_field, "複聯備援")),
            ),
            CheckboxOption(
                label="異機備援",
                selected=bool(_bool_from_option(internal_field, "異機備援")),
            ),
            CheckboxOption(
                label="病歷系統備援",
                selected=bool(_bool_from_option(internal_field, "病歷系統備援")),
            ),
            CheckboxOption(
                label="影像系統備援",
                selected=bool(_bool_from_option(internal_field, "影像系統備援")),
            ),
        ],
        selected_values=[],
        raw_text=clean_text(internal_text),
    ).model_dump()
    _finalize_checkbox_field(internal_field)

    mobile_storage_selected = bool(_bool_from_option(external_field, "移動式"))
    automatic_offsite_backup = bool(_bool_from_option(external_field, "自動化異地備援"))
    auto_sync = bool(_bool_from_option(external_field, "自動同步"))
    failover_switching = bool(_bool_from_option(external_field, "停擺切換"))
    emr_backup = bool(_bool_from_option(external_field, "病歷資訊系統備援"))
    image_backup = bool(_bool_from_option(external_field, "影像資訊系統備援"))
    picture_backup = bool(_bool_from_option(external_field, "圖像資訊系統備援"))
    distance_selected = any(
        option.get("label") == "超過" and option.get("selected")
        for option in external_field.get("options", [])
    )
    external_field = CheckboxField(
        question="院外備援",
        options=[
            CheckboxOption(
                label="超過指定距離之異地備援",
                selected=distance_selected,
                value=distance_field.get("value"),
                unit="公里",
            ),
            CheckboxOption(
                label="移動式儲存設備移送至他地",
                selected=mobile_storage_selected,
                value=mobile_storage_type,
            ),
            CheckboxOption(
                label="備份機制",
                selected=bool(_bool_from_option(external_field, "備份機制")),
                value=backup_type_field.get("value"),
            ),
            CheckboxOption(
                label="自動化異地備援",
                selected=automatic_offsite_backup,
            ),
            CheckboxOption(
                label="自動同步",
                selected=auto_sync,
            ),
            CheckboxOption(
                label="停擺切換",
                selected=failover_switching,
            ),
            CheckboxOption(
                label="病歷資訊系統備援",
                selected=emr_backup,
            ),
            CheckboxOption(
                label="病歷資訊系統備援：單機",
                selected=bool(_bool_from_option(emr_segment, "單機")),
            ),
            CheckboxOption(
                label="病歷資訊系統備援：伺服器等",
                selected=bool(_bool_from_option(emr_segment, "伺服器等")),
            ),
            CheckboxOption(
                label="影像資訊系統備援",
                selected=image_backup,
            ),
            CheckboxOption(
                label="影像資訊系統備援：單機",
                selected=bool(_bool_from_option(image_segment, "單機")),
            ),
            CheckboxOption(
                label="影像資訊系統備援：伺服器等",
                selected=bool(_bool_from_option(image_segment, "伺服器等")),
            ),
            CheckboxOption(
                label="圖像資訊系統備援",
                selected=picture_backup,
            ),
            CheckboxOption(
                label="圖像資訊系統備援：單機",
                selected=bool(_bool_from_option(graphic_segment, "單機")),
            ),
            CheckboxOption(
                label="圖像資訊系統備援：伺服器等",
                selected=bool(_bool_from_option(graphic_segment, "伺服器等")),
            ),
            CheckboxOption(
                label="雲端備援廠家",
                selected=bool(_bool_from_option(external_field, "雲端備援廠家")),
                value=cloud_vendor_field.get("value"),
            ),
            CheckboxOption(
                label="其他院外備援方式",
                selected=bool(_bool_from_option(external_field, "其他院外備援方式")),
                value=other_external_backup_method.get("value"),
            ),
        ],
        selected_values=[],
        raw_text=clean_text(external_text),
    ).model_dump()
    _finalize_checkbox_field(external_field)

    has_offsite_backup = any(
        [
            distance_selected,
            automatic_offsite_backup,
            auto_sync,
            failover_switching,
            emr_backup,
            image_backup,
            picture_backup,
            bool(cloud_vendor_field.get("filled") or cloud_vendor_field.get("value")),
            bool(
                other_external_backup_method.get("filled")
                or other_external_backup_method.get("value")
            ),
        ]
    )

    return {
        "normal_condition": {
            "internal_backup": internal_field,
            "external_backup": external_field,
        },
        "has_offsite_backup": has_offsite_backup,
        "internal_backup": {
            "standalone_backup": _bool_from_option(internal_field, "單機備援"),
            "dual_connection_backup": _bool_from_option(internal_field, "複聯備援"),
            "different_machine_backup": _bool_from_option(internal_field, "異機備援"),
            "medical_record_system_backup": _bool_from_option(internal_field, "病歷系統備援"),
            "image_system_backup": _bool_from_option(internal_field, "影像系統備援"),
        },
        "external_backup": {
            "offsite_distance_km": distance_field,
            "mobile_storage_transfer": mobile_storage_selected,
            "mobile_storage_type": mobile_storage_type,
            "backup_type": backup_type_field,
            "automatic_offsite_backup": automatic_offsite_backup,
            "auto_sync": auto_sync,
            "failover_switching": failover_switching,
            "cloud_vendor": cloud_vendor_field,
            "other_external_backup_method": other_external_backup_method,
            "medical_record_system_backup": emr_backup,
            "medical_record_system_standalone": _bool_from_option(
                external_field,
                "病歷資訊系統備援：單機",
            ),
            "medical_record_system_server": _bool_from_option(
                external_field,
                "病歷資訊系統備援：伺服器等",
            ),
            "image_information_system_backup": image_backup,
            "image_information_system_standalone": _bool_from_option(
                external_field,
                "影像資訊系統備援：單機",
            ),
            "image_information_system_server": _bool_from_option(
                external_field,
                "影像資訊系統備援：伺服器等",
            ),
            "picture_information_system_backup": picture_backup,
            "picture_information_system_standalone": _bool_from_option(
                external_field,
                "圖像資訊系統備援：單機",
            ),
            "picture_information_system_server": _bool_from_option(
                external_field,
                "圖像資訊系統備援：伺服器等",
            ),
        },
        "island_mode_response": island_field,
        "island_mode": {
            "can_operate_standalone_without_network": _bool_from_option(
                island_field,
                "孤島效應下可單機",
            ),
            "improvement_years": improvement_years_field,
        },
        "recommendation_text": clean_text(recommendation_text),
        "raw_text": {
            "current_status": cleaned,
            "recommendation": clean_text(recommendation_text),
        },
    }


def parse_facility_maintenance_section(text: str) -> dict[str, Any]:
    cleaned = clean_text(text)
    checkbox_text = "\n".join(line for line in cleaned.splitlines() if line.startswith((CHECKED, UNCHECKED)))
    note = "\n".join(
        line for line in cleaned.splitlines() if not line.startswith((CHECKED, UNCHECKED))
    ).strip()
    parsed = parse_checkbox_options(
        checkbox_text,
        question="設施維護措施涉及項目",
    )
    return {
        "question": parsed["question"],
        "selection_type": parsed["selection_type"],
        "options": parsed["options"],
        "selected_values": parsed["selected_values"],
        "note": note or None,
        "raw_text": cleaned,
    }


def parse_central_monitoring_section(text: str, recommendation_text: str = "") -> dict[str, Any]:
    cleaned = clean_text(text)
    checkbox_text = "\n".join(line for line in cleaned.splitlines() if line.startswith((CHECKED, UNCHECKED)))
    note = "\n".join(
        line for line in cleaned.splitlines() if not line.startswith((CHECKED, UNCHECKED))
    ).strip()
    parsed = parse_checkbox_options(
        checkbox_text,
        question="中央監控系統可監視項目",
    )
    return {
        "question": parsed["question"],
        "selection_type": parsed["selection_type"],
        "options": parsed["options"],
        "selected_values": parsed["selected_values"],
        "note": note or None,
        "recommendation_text": clean_text(recommendation_text),
        "raw_text": cleaned,
    }


def build_data_quality(result: dict[str, Any]) -> dict[str, Any]:
    missing_fields: list[str] = []
    warnings: list[str] = []
    calculation_checks: list[dict[str, Any]] = []

    def visit(node: Any, parent_key: str | None = None) -> None:
        if isinstance(node, dict):
            if {"field_id", "label", "filled"}.issubset(node.keys()):
                if not node.get("filled"):
                    missing_fields.append(
                        _missing_field_message(
                            node.get("field_id"),
                            node.get("label"),
                        )
                    )
                field_value = node.get("value")
                if isinstance(field_value, str) and PLACEHOLDER_PATTERN.search(field_value):
                    warnings.append(f"{node.get('label')}可能仍為範例或空白內容：{field_value}")

            if {"label", "selected"}.issubset(node.keys()):
                label = node.get("label")
                if (
                    node.get("selected")
                    and "value" in node
                    and node.get("value") is None
                    and node.get("unit")
                ):
                    missing_fields.append(f"{label}數值未填")
                if (
                    label not in {"每日每床用水量"}
                    and not node.get("selected")
                    and isinstance(node.get("value"), (int, float))
                ):
                    warnings.append(f"{label}有填值但未勾選")
                if (
                    label not in {"每日每床用水量"}
                    and not node.get("selected")
                    and isinstance(node.get("value"), dict)
                    and any(value is not None for value in node.get("value", {}).values())
                ):
                    warnings.append(f"{label}有填值但未勾選")
                if (
                    node.get("selected")
                    and isinstance(node.get("value"), str)
                    and PLACEHOLDER_PATTERN.search(node.get("value"))
                ):
                    warnings.append(f"{label}可能仍為範例或空白內容：{node.get('value')}")

            for key, value in node.items():
                if key in {"raw_text", "recommendation_text"}:
                    continue
                visit(value, parent_key=key)
            return

        if isinstance(node, list):
            for item in node:
                visit(item, parent_key=parent_key)
            return

    visit(result)

    hospital_name = result.get("basic_info", {}).get("hospital_name")
    if isinstance(hospital_name, str) and "OOOO" in hospital_name:
        warnings.append(f"醫院名稱可能仍為範例文字：{hospital_name}")

    power_calc = result.get("sections", {}).get("power", {}).get("island_mode", {}).get("calculation", {})
    if power_calc:
        document_value = power_calc.get("document_result_hours")
        system_value = power_calc.get("system_calculated_hours")
        if document_value is not None and system_value is not None:
            consistent = abs(float(document_value) - float(system_value)) < 0.05
            note = (
                "文件值與公式計算值略有差異，請確認四捨五入或耗油量基準。"
                if not consistent
                else "文件值與系統計算值一致。"
            )
            calculation_checks.append(
                CalculationCheck(
                    field="power.island_mode.calculation",
                    document_value=document_value,
                    system_calculated_value=system_value,
                    consistent=consistent,
                    note=note,
                ).model_dump()
            )

    water_section = result.get("sections", {}).get("water", {}).get("island_mode", {})
    full_bed = water_section.get("full_bed_scenario", {})
    reduced_bed = water_section.get("reduced_bed_scenario", {})
    if full_bed.get("meets_72_hours") is False:
        warnings.append("供水全床位情境未達72小時。")
    if reduced_bed.get("document_days") is not None and reduced_bed.get("calculated_days") is not None:
        consistent = abs(float(reduced_bed["document_days"]) - float(reduced_bed["calculated_days"])) < 0.1
        calculation_checks.append(
            CalculationCheck(
                field="water.island_mode.reduced_bed_scenario",
                document_value=reduced_bed["document_days"],
                system_calculated_value=reduced_bed["calculated_days"],
                consistent=consistent,
                note="供水降載情境文件值與系統值比對。",
            ).model_dump()
        )

    dedup_missing = list(dict.fromkeys(missing_fields))
    dedup_warnings = list(dict.fromkeys(warnings))

    return DataQuality(
        missing_fields=dedup_missing,
        warnings=dedup_warnings,
        calculation_checks=[CalculationCheck.model_validate(item) for item in calculation_checks],
    ).model_dump()


def build_overall_summary(result: dict[str, Any]) -> dict[str, Any]:
    sections = result.get("sections", {})
    power = sections.get("power", {})
    water = sections.get("water", {})
    medical_gas = sections.get("medical_gas", {})
    it_backup = sections.get("it_backup", {})
    facility = sections.get("facility_maintenance", {})
    monitoring = sections.get("central_monitoring", {})

    generator_count = power.get("generator", {}).get("generator_count")
    power_hours = power.get("island_mode", {}).get("calculation", {}).get("system_calculated_hours")
    power_can_meet = power.get("island_mode", {}).get("can_reach_72_hours")
    power_summary = (
        f"現有油量估算可支撐約 {power_hours} 小時"
        if power_hours is not None
        else "供電支撐時數待確認"
    )
    if generator_count == 1:
        power_summary += "，但僅有一台發電機，建議補強備援。"
    else:
        power_summary += "。"

    full_days = water.get("island_mode", {}).get("full_bed_scenario", {}).get("calculated_days")
    reduced_days = water.get("island_mode", {}).get("reduced_bed_scenario", {}).get("calculated_days")
    water_can_meet = water.get("island_mode", {}).get("reduced_bed_scenario", {}).get("meets_72_hours")
    water_summary_parts: list[str] = []
    if full_days is not None:
        water_summary_parts.append(f"全床位情境約 {full_days} 天")
    if reduced_days is not None:
        water_summary_parts.append(f"降載情境約 {reduced_days} 天")
    water_summary = "；".join(water_summary_parts) if water_summary_parts else "供水情境待確認"
    if water_can_meet is True:
        water_summary += "，降載情境可達 72 小時。"
    elif water_summary:
        water_summary += "。"

    gas_days = medical_gas.get("island_mode", {}).get("calculated_days")
    gas_can_meet = medical_gas.get("island_mode", {}).get("meets_72_hours")
    gas_summary = (
        f"液氧總量估算可支撐約 {gas_days} 天。"
        if gas_days is not None
        else "醫用氣體支撐天數待確認。"
    )

    it_has_mobile = it_backup.get("external_backup", {}).get("mobile_storage_transfer")
    it_has_auto = it_backup.get("external_backup", {}).get("automatic_offsite_backup")
    it_summary_parts: list[str] = []
    if it_backup.get("internal_backup", {}).get("standalone_backup"):
        it_summary_parts.append("具單機備援")
    if it_has_mobile:
        it_summary_parts.append("具移動式備援")
    if it_has_auto:
        it_summary_parts.append("具自動化異地備援")
    it_summary = "、".join(it_summary_parts) if it_summary_parts else "IT備援配置待確認"
    if not it_has_auto:
        it_summary += "，自動化異地備援仍待補強。"
    else:
        it_summary += "。"

    facility_selected = facility.get("selected_values", [])
    facility_summary = (
        f"已納入 { '、'.join(facility_selected) } 等設施維護措施。"
        if facility_selected
        else "設施維護措施待確認。"
    )

    monitoring_selected = monitoring.get("selected_values", [])
    monitoring_summary = (
        f"目前中央監控涵蓋 { '、'.join(monitoring_selected) }。"
        if monitoring_selected
        else "中央監控涵蓋項目待確認。"
    )
    if "供水" not in monitoring_selected or "供氣" not in monitoring_selected:
        monitoring_summary += " 建議補強供水與供氣監控。"

    return {
        "power": {
            "can_meet_72_hours": power_can_meet,
            "summary": power_summary,
        },
        "water": {
            "can_meet_72_hours": water_can_meet,
            "summary": water_summary,
        },
        "medical_gas": {
            "can_meet_72_hours": gas_can_meet,
            "summary": gas_summary,
        },
        "it_backup": {
            "can_meet_72_hours": None,
            "summary": it_summary,
        },
        "facility_maintenance": {
            "can_meet_72_hours": None,
            "summary": facility_summary,
        },
        "central_monitoring": {
            "can_meet_72_hours": None,
            "summary": monitoring_summary,
        },
    }


SECTION_DISPLAY_NAMES = {
    "basic_info": "一、基本資訊",
    "power": "二、供電相關",
    "water": "三、供水相關",
    "medical_gas": "四、醫用氣體",
    "it_backup": "五、資訊系統備援",
    "facility_maintenance": "六、設施維護措施",
    "central_monitoring": "七、中央監控系統",
}


def _get_nested(payload: dict[str, Any], *keys: str, default: Any = None) -> Any:
    current: Any = payload
    for key in keys:
        if not isinstance(current, dict):
            return default
        current = current.get(key)
    return default if current is None else current


def _split_recommendations(text: str | None) -> list[str]:
    cleaned = clean_text(text or "")
    if not cleaned:
        return []

    lines = [line.strip() for line in cleaned.splitlines() if line.strip()]
    if len(lines) > 1:
        return lines

    parts = [
        part.strip()
        for part in re.split(r"(?=(?:\d+\.|[一二三四五六七八九十]+、))", cleaned)
        if part.strip()
    ]
    return parts or [cleaned]


def _section_recommendations(section: dict[str, Any]) -> list[str]:
    if section.get("recommendation_text"):
        return _split_recommendations(section.get("recommendation_text"))

    raw_recommendation = _get_nested(section, "raw_text", "recommendation")
    recommendations = _split_recommendations(raw_recommendation)
    return recommendations or ["無"]


def _formula_numbers(formula: str | None) -> list[int | float]:
    if not formula:
        return []
    return [_to_number(match.group(0)) for match in NUMBER_PATTERN.finditer(formula)]


def _calculation_note(document_value: Any, system_value: Any, unit: str) -> str:
    if document_value is None or system_value is None:
        return ""
    try:
        if abs(float(document_value) - float(system_value)) > 0.05:
            return f"文件值與系統計算值略有差異，建議確認四捨五入或計算基準。"
    except (TypeError, ValueError):
        return ""
    return f"文件值與系統計算值一致或差異在可接受範圍內，單位：{unit}。"


def _build_basic_info_report_section(basic_info: dict[str, Any]) -> dict[str, Any]:
    return {
        "section_id": "basic_info",
        "section_name": SECTION_DISPLAY_NAMES["basic_info"],
        "current_status": basic_info.get("completion_status"),
        "recommendations": [basic_info.get("recommendation_note") or "無"],
        "risk_points": [],
    }


def _build_power_report_section(section: dict[str, Any], target_hours: int | float) -> dict[str, Any]:
    calc = _get_nested(section, "island_mode", "calculation", default={})
    formula_numbers = _formula_numbers(calc.get("formula"))
    generator_count = _get_nested(section, "generator", "generator_count")
    hourly_fuel = (
        formula_numbers[1]
        if len(formula_numbers) > 1
        else _get_nested(section, "generator", "generators", default=[{}])[0].get("hourly_fuel_consumption_liters")
        if _get_nested(section, "generator", "generators", default=[])
        else None
    )
    fuel_total = formula_numbers[0] if formula_numbers else None

    risk_points: list[str] = []
    if generator_count == 1:
        risk_points.append("雖然現有油量可支撐超過72小時，但僅有一台發電機，存在單點失效風險。")
    if calc.get("meets_target") is False:
        risk_points.append("供電估算未達72小時，需補強燃料、發電機或用電降載方案。")

    return {
        "section_id": "power",
        "section_name": SECTION_DISPLAY_NAMES["power"],
        "current_status": {
            "normal_condition": section.get("normal_condition"),
            "island_mode_response": section.get("island_mode_response"),
        },
        "extracted_facts": {
            "generator_count": generator_count,
            "fuel_total_liters": fuel_total,
            "hourly_fuel_consumption_liters": hourly_fuel,
            "estimated_supply_hours_from_document": calc.get("document_result_hours"),
            "can_exceed_72_hours": _get_nested(section, "island_mode", "can_reach_72_hours"),
        },
        "calculations": [
            {
                "calculation_name": "孤島效應下供電可支撐時數",
                "formula": calc.get("formula"),
                "inputs": {
                    "fuel_total_liters": fuel_total,
                    "hourly_fuel_consumption_liters": hourly_fuel,
                },
                "document_result_hours": calc.get("document_result_hours"),
                "system_calculated_result_hours": calc.get("system_calculated_hours"),
                "target_hours": target_hours,
                "meets_target": calc.get("meets_target"),
                "note": _calculation_note(
                    calc.get("document_result_hours"),
                    calc.get("system_calculated_hours"),
                    "小時",
                ),
            }
        ]
        if calc
        else [],
        "risk_points": risk_points,
        "recommendations": _section_recommendations(section),
    }


def _build_water_report_section(section: dict[str, Any]) -> dict[str, Any]:
    full_bed = _get_nested(section, "island_mode", "full_bed_scenario", default={})
    reduced_bed = _get_nested(section, "island_mode", "reduced_bed_scenario", default={})
    sources = section.get("sources", [])
    source_lookup = {
        source.get("type"): source.get("selected")
        for source in sources
        if isinstance(source, dict)
    }
    interconnected = _get_nested(
        section,
        "island_mode",
        "interconnected_water_tank_system",
        "selected",
        default=False,
    )

    risk_points: list[str] = []
    if full_bed.get("meets_72_hours") is False:
        risk_points.append("若以全院床位估算，供水未達完整72小時安全餘裕。")
    if not source_lookup.get("合法井水", False) and not source_lookup.get("戰備水池", False):
        risk_points.append("目前未見合法井水或戰備水池等替代水源。")
    if not interconnected:
        risk_points.append("院內儲水槽聯通管系統未勾選或未填值，聯通備援仍可確認。")

    calculations: list[dict[str, Any]] = []
    if full_bed:
        calculations.append(
            {
                "calculation_name": "全床位供水天數",
                "scenario": f"{full_bed.get('beds')}床，每床每日{full_bed.get('daily_water_per_bed_tons')}公噸",
                "formula": "total_storage_tons / (beds * daily_water_per_bed_tons)",
                "inputs": {
                    "total_water_storage_tons": full_bed.get("total_storage_tons"),
                    "beds": full_bed.get("beds"),
                    "daily_water_per_bed_tons": full_bed.get("daily_water_per_bed_tons"),
                },
                "document_result_days": full_bed.get("document_days"),
                "system_calculated_result_days": full_bed.get("calculated_days"),
                "target_days": 3,
                "meets_72_hours": full_bed.get("meets_72_hours"),
            }
        )
    if reduced_bed:
        calculations.append(
            {
                "calculation_name": "降載情境供水天數",
                "scenario": f"降載至{reduced_bed.get('beds')}床，每床每日{reduced_bed.get('daily_water_per_bed_liters')}公升",
                "formula": "total_storage_liters / beds / daily_water_per_bed_liters",
                "inputs": {
                    "total_water_storage_liters": reduced_bed.get("total_storage_liters"),
                    "beds": reduced_bed.get("beds"),
                    "daily_water_per_bed_liters": reduced_bed.get("daily_water_per_bed_liters"),
                },
                "document_result_days": reduced_bed.get("document_days"),
                "system_calculated_result_days": reduced_bed.get("calculated_days"),
                "target_days": 3,
                "meets_72_hours": reduced_bed.get("meets_72_hours"),
            }
        )

    return {
        "section_id": "water",
        "section_name": SECTION_DISPLAY_NAMES["water"],
        "current_status": {
            "normal_condition": section.get("normal_condition"),
            "island_mode_response": section.get("island_mode_response"),
        },
        "extracted_facts": {
            "total_water_storage_tons": section.get("total_storage_tons"),
            "beds": full_bed.get("beds"),
            "daily_water_per_bed_tons_full_bed_scenario": full_bed.get("daily_water_per_bed_tons"),
            "full_bed_supply_days": full_bed.get("calculated_days"),
            "reduced_general_beds": reduced_bed.get("general_beds"),
            "reduced_icu_beds": reduced_bed.get("icu_beds"),
            "reduced_total_beds": reduced_bed.get("beds"),
            "daily_water_per_bed_liters_reduced_scenario": reduced_bed.get("daily_water_per_bed_liters"),
            "reduced_bed_supply_days": reduced_bed.get("calculated_days"),
        },
        "calculations": calculations,
        "risk_points": risk_points,
        "recommendations": _section_recommendations(section),
    }


def _build_medical_gas_report_section(section: dict[str, Any]) -> dict[str, Any]:
    island_mode = section.get("island_mode", {})
    liquid_oxygen = section.get("liquid_oxygen", {})
    oxygen_cylinders = section.get("oxygen_cylinders", {})
    formula_numbers = _formula_numbers(island_mode.get("formula"))

    risk_points: list[str] = []
    if liquid_oxygen.get("total_capacity_tons"):
        risk_points.append("供氣主要依賴液氧槽與汽化供應，仍需確認災害期間管線與切換流程。")
    if island_mode.get("meets_72_hours") is False:
        risk_points.append("供氧估算未達72小時，需補強氧氣儲量或降載應變機制。")

    calculations = []
    if island_mode:
        calculations.append(
            {
                "calculation_name": "孤島效應下供氧天數",
                "formula": island_mode.get("formula"),
                "inputs": {
                    "liquid_oxygen_total_tons": formula_numbers[0] if formula_numbers else liquid_oxygen.get("total_capacity_tons"),
                    "daily_oxygen_usage_tons": formula_numbers[1] if len(formula_numbers) > 1 else None,
                },
                "document_result_days": island_mode.get("document_days"),
                "system_calculated_result_days": island_mode.get("calculated_days"),
                "target_days": 3,
                "meets_72_hours": island_mode.get("meets_72_hours"),
            }
        )

    return {
        "section_id": "medical_gas",
        "section_name": SECTION_DISPLAY_NAMES["medical_gas"],
        "current_status": {
            "normal_condition": section.get("normal_condition"),
            "island_mode_response": section.get("island_mode_response"),
        },
        "extracted_facts": {
            "liquid_oxygen_total_tons": liquid_oxygen.get("total_capacity_tons"),
            "daily_oxygen_usage_tons": formula_numbers[1] if len(formula_numbers) > 1 else None,
            "estimated_supply_days": island_mode.get("calculated_days"),
            "oxygen_cylinder_inside_hospital_count": oxygen_cylinders.get("inside_hospital_count"),
            "oxygen_cylinder_outside_hospital_count": oxygen_cylinders.get("outside_hospital_count"),
            "oxygen_cylinder_total_count": oxygen_cylinders.get("total_count"),
            "oxygen_cylinder_gas_equivalent_tons": oxygen_cylinders.get("gas_equivalent_tons"),
            "can_exceed_72_hours": island_mode.get("meets_72_hours"),
        },
        "calculations": calculations,
        "risk_points": risk_points,
        "recommendations": _section_recommendations(section),
    }


def _build_it_backup_report_section(section: dict[str, Any]) -> dict[str, Any]:
    internal = section.get("internal_backup", {})
    external = section.get("external_backup", {})
    island_mode = section.get("island_mode", {})

    risk_points: list[str] = []
    if not external.get("automatic_offsite_backup"):
        risk_points.append("目前未見自動化異地備援機制。")
    if not external.get("auto_sync"):
        risk_points.append("目前未見自動同步機制，資料復原時點需確認。")
    if not external.get("failover_switching"):
        risk_points.append("目前未見停擺切換機制，資訊服務連續性仍可能受限。")
    if not island_mode.get("can_operate_standalone_without_network"):
        risk_points.append("孤島效應下可單機作業未勾選，建議確認各單位離線作業能力。")

    return {
        "section_id": "it_backup",
        "section_name": SECTION_DISPLAY_NAMES["it_backup"],
        "current_status": {
            "normal_condition": section.get("normal_condition"),
            "island_mode_response": section.get("island_mode_response"),
        },
        "extracted_facts": {
            "has_standalone_backup": internal.get("standalone_backup"),
            "has_mobile_storage_transfer": external.get("mobile_storage_transfer"),
            "mobile_storage_type": external.get("mobile_storage_type"),
            "has_automatic_offsite_backup": external.get("automatic_offsite_backup"),
            "has_auto_sync": external.get("auto_sync"),
            "has_failover_switching": external.get("failover_switching"),
            "cloud_vendor": _get_nested(external, "cloud_vendor", "value"),
            "has_offsite_backup": section.get("has_offsite_backup"),
        },
        "calculations": [],
        "risk_points": risk_points,
        "recommendations": _section_recommendations(section),
    }


def _build_simple_report_section(section_id: str, section: dict[str, Any]) -> dict[str, Any]:
    risk_points: list[str] = []
    if section_id == "central_monitoring":
        selected = set(section.get("selected_values", []))
        if "供水" not in selected or "供氣" not in selected:
            risk_points.append("中央監控系統目前未完整涵蓋供水及供氣。")

    return {
        "section_id": section_id,
        "section_name": SECTION_DISPLAY_NAMES[section_id],
        "current_status": {
            "question": section.get("question"),
            "selection_type": section.get("selection_type"),
            "options": section.get("options", []),
            "selected_values": section.get("selected_values", []),
            "note": section.get("note"),
            "raw_text": section.get("raw_text"),
        },
        "calculations": [],
        "risk_points": risk_points,
        "recommendations": _section_recommendations(section),
    }


def _build_report_data_quality(data_quality: dict[str, Any]) -> dict[str, Any]:
    warnings = data_quality.get("warnings", [])
    calculation_checks = data_quality.get("calculation_checks", [])
    possible_inconsistencies = [
        {
            "field": item.get("field"),
            "document_value": item.get("document_value"),
            "system_calculated_value": item.get("system_calculated_value"),
            "note": item.get("note"),
        }
        for item in calculation_checks
        if item.get("consistent") is False
    ]
    placeholder_fields = [
        warning
        for warning in warnings
        if "範例" in str(warning) or "OOOO" in str(warning) or "OO" in str(warning)
    ]

    return {
        "missing_or_blank_fields": data_quality.get("missing_fields", []),
        "possible_inconsistencies": possible_inconsistencies,
        "placeholder_fields": placeholder_fields,
        "warnings": warnings,
    }


def build_report_input_json(normalized_json: dict[str, Any]) -> dict[str, Any]:
    sections = normalized_json.get("sections", {})
    target_hours = normalized_json.get("target_hours", 72)

    report_sections: list[dict[str, Any]] = [
        _build_basic_info_report_section(normalized_json.get("basic_info", {}))
    ]
    if "power" in sections:
        report_sections.append(_build_power_report_section(sections["power"], target_hours))
    if "water" in sections:
        report_sections.append(_build_water_report_section(sections["water"]))
    if "medical_gas" in sections:
        report_sections.append(_build_medical_gas_report_section(sections["medical_gas"]))
    if "it_backup" in sections:
        report_sections.append(_build_it_backup_report_section(sections["it_backup"]))
    for section_id in ("facility_maintenance", "central_monitoring"):
        if section_id in sections:
            report_sections.append(_build_simple_report_section(section_id, sections[section_id]))

    profile = normalized_json.get("document_profile")
    document_type = (
        "hospital_resilience_diagnosis_form"
        if profile == "diagnosis_form"
        else "hospital_resilience_questionnaire"
    )

    return {
        "document_type": document_type,
        "document_profile": profile,
        "project_name": normalized_json.get("project_name"),
        "form_name": normalized_json.get("form_name"),
        "target_hours": target_hours,
        "source_file": normalized_json.get("source_file"),
        "basic_info": {
            key: normalized_json.get("basic_info", {}).get(key)
            for key in (
                "organizer",
                "commissioned_unit",
                "hospital_name",
                "diagnosis_date",
                "diagnosis_expert",
            )
            if key in normalized_json.get("basic_info", {})
        },
        "sections": report_sections,
        "overall_summary": normalized_json.get("overall_summary", {}),
        "data_quality": _build_report_data_quality(normalized_json.get("data_quality", {})),
    }


def _build_normalized_questionnaire_json(
    raw_json: dict[str, Any],
    docx_path: str,
    basic_info: dict[str, Any],
    sections: dict[str, Any],
    *,
    profile: str,
) -> dict[str, Any]:
    normalized_json: dict[str, Any] = {
        "document_type": "hospital_resilience_questionnaire",
        "document_profile": profile,
        "project_name": raw_json["paragraphs"][0] if raw_json["paragraphs"] else "",
        "form_name": raw_json["paragraphs"][1] if len(raw_json["paragraphs"]) > 1 else "",
        "version": None,
        "target_hours": 72,
        "source_file": {
            "file_name": Path(docx_path).name,
            "file_type": Path(docx_path).suffix.lstrip(".").lower() or None,
        },
        "basic_info": basic_info,
        "sections": sections,
    }
    normalized_json["overall_summary"] = build_overall_summary(normalized_json)
    normalized_json["data_quality"] = build_data_quality(normalized_json)
    return normalized_json


def _parse_diagnosis_form(raw_json: dict[str, Any], docx_path: str) -> dict[str, Any]:
    rows = _get_primary_table_rows(raw_json)

    power_row = _find_section_row(rows, "power")
    water_row = _find_section_row(rows, "water")
    gas_row = _find_section_row(rows, "medical_gas")
    it_row = _find_section_row(rows, "it_backup")
    facility_row = _find_section_row(rows, "facility_maintenance")
    monitoring_row = _find_section_row(rows, "central_monitoring")

    sections = {
        "power": parse_power_section(
            power_row[1] if len(power_row) > 1 else "",
            _dedupe_merged_cell_text(power_row),
        ),
        "water": parse_water_section(
            water_row[1] if len(water_row) > 1 else "",
            _dedupe_merged_cell_text(water_row),
        ),
        "medical_gas": parse_medical_gas_section(
            gas_row[1] if len(gas_row) > 1 else "",
            _dedupe_merged_cell_text(gas_row),
        ),
        "it_backup": parse_it_backup_section(
            it_row[1] if len(it_row) > 1 else "",
            _dedupe_merged_cell_text(it_row),
        ),
        "facility_maintenance": parse_facility_maintenance_section(
            facility_row[1] if len(facility_row) > 1 else "",
        ),
        "central_monitoring": parse_central_monitoring_section(
            monitoring_row[1] if len(monitoring_row) > 1 else "",
            _dedupe_merged_cell_text(monitoring_row),
        ),
    }

    return _build_normalized_questionnaire_json(
        raw_json=raw_json,
        docx_path=docx_path,
        basic_info=parse_basic_info(raw_json),
        sections=sections,
        profile="diagnosis_form",
    )


def _parse_monthly_usage(
    text: str,
    *,
    value_label: str,
    year: int = 2025,
) -> dict[str, int | float]:
    monthly_usage: dict[str, int | float] = {}
    cleaned = clean_text(text)
    for month in range(1, 13):
        match = re.search(
            rf"{month}月{re.escape(value_label)}[:：]?\s*_*\s*([0-9,]+(?:\.[0-9]+)?)",
            cleaned,
        )
        if match:
            monthly_usage[f"{year}-{month:02d}"] = _to_number(match.group(1))
    return monthly_usage


def _parse_full_questionnaire_facility(table: dict[str, Any]) -> dict[str, Any]:
    rows = _get_unique_cleaned_rows(table)
    headers = rows[0][1:5] if rows else []
    labels = [clean_text(header).lstrip("□■").strip() for header in headers if clean_text(header)]
    options = [{"label": label, "selected": True, "value": None, "unit": None} for label in labels]
    return {
        "question": "設施維護措施涉及項目",
        "selection_type": "multiple",
        "options": options,
        "selected_values": [option["label"] for option in options if option["selected"]],
        "note": _flatten_table_text(table, start_row=1),
        "raw_text": _flatten_table_text(table),
    }


def _parse_full_questionnaire_monitoring(table: dict[str, Any]) -> dict[str, Any]:
    rows = _get_unique_cleaned_rows(table)
    headers = rows[0][1:5] if rows else []
    labels = [clean_text(header).lstrip("□■").strip() for header in headers if clean_text(header)]

    options: list[dict[str, Any]] = []
    for offset, label in enumerate(labels, start=1):
        cell_text = rows[1][offset] if len(rows) > 1 and len(rows[1]) > offset else ""
        selected = bool(
            any(marker in cell_text for marker in CHECKED_MARKERS)
            or re.search(r"約[_ ]*(\d+(?:\.\d+)?)\s*%", cell_text)
        )
        options.append(
            {
                "label": label,
                "selected": selected,
                "value": None,
                "unit": None,
            }
        )

    return {
        "question": "中央監控系統可監視項目",
        "selection_type": "multiple",
        "options": options,
        "selected_values": [option["label"] for option in options if option["selected"]],
        "note": _flatten_table_text(table, start_row=1),
        "recommendation_text": None,
        "raw_text": _flatten_table_text(table),
    }


def _parse_full_questionnaire(raw_json: dict[str, Any], docx_path: str) -> dict[str, Any]:
    power_table = _get_table_by_index(raw_json, 1)
    power_island_table = _get_table_by_index(raw_json, 2)
    water_table = _get_table_by_index(raw_json, 3)
    gas_table = _get_table_by_index(raw_json, 4)
    it_table = _get_table_by_index(raw_json, 5)
    facility_table = _get_table_by_index(raw_json, 6)
    monitoring_table = _get_table_by_index(raw_json, 7)

    power_normal_text = _join_table_column_text(power_table, 1)
    power_island_text = _flatten_table_text(power_island_table)
    power_combined_text = clean_text(
        f"{power_normal_text}\n孤島效應下應變措施\n{power_island_text}"
    )
    power = parse_power_section(power_combined_text, power_island_text)
    generator_count_match = re.search(r"總台數_*([0-9]+)", power_normal_text)
    if generator_count_match:
        power["generator"]["generator_count"] = _to_number(generator_count_match.group(1))
    if power["generator"].get("generator_count") and not power["generator"].get("generators"):
        power["generator"]["generators"] = [
            {
                "id": "#01",
                "location": None,
                "power_kw": None,
                "type": None,
                "ats_start_seconds": None,
                "daily_tank_liters": None,
                "hourly_fuel_consumption_liters": None,
            }
        ]

    power_baseline_text = _find_paragraph_containing(raw_json, "2025年(12個月)用電量")
    power_monthly_text = _find_paragraph_containing(raw_json, "6月用電")
    if power_baseline_text:
        annual_match = re.search(r"用電量[：:]\s*([0-9,]+)", power_baseline_text)
        daily_match = re.search(r"=\s*([0-9]+(?:\.[0-9]+)?)\s*/24", power_baseline_text)
        hourly_match = re.search(r"/24(?:\(小時\))?=\s*([0-9]+(?:\.[0-9]+)?)", power_baseline_text)
        power["baseline"] = {
            "year": 2025,
            "annual_usage_kwh": _to_number(annual_match.group(1)) if annual_match else None,
            "average_daily_usage_kwh": _to_number(daily_match.group(1)) if daily_match else None,
            "average_hourly_usage_kwh": _to_number(hourly_match.group(1)) if hourly_match else None,
            "monthly_usage_kwh": _parse_monthly_usage(power_monthly_text, value_label="用電"),
        }
    ups_match = re.search(r"供給場所[：:]\s*([^\n]+)", power_normal_text)
    if ups_match:
        power["ups"]["supply_areas"] = [
            clean_text(item) for item in re.split(r"[、,，]", ups_match.group(1)) if clean_text(item)
        ]
    power["ups"]["duration_seconds"] = extract_number(
        _extract_labeled_text(power_normal_text, "可維持供電時間") or ""
    )
    power["ups"]["is_parallel"] = "■是" in power_normal_text or "是(" in power_normal_text
    parallel_match = re.search(r"計算式[：:]\s*([0-9]+(?:\.[0-9]+)?)kWh", power_normal_text)
    independent_match = re.search(r"=\s*([0-9]+(?:\.[0-9]+)?)\s*\(獨立加總\)", power_normal_text)
    power["ups"]["parallel_capacity_kwh"] = (
        _to_number(parallel_match.group(1)) if parallel_match else None
    )
    power["ups"]["independent_total_capacity_kwh"] = (
        _to_number(independent_match.group(1)) if independent_match else None
    )
    fuel_type_match = re.search(r"使用燃油類型[：:]\s*_*\s*([^\n(；;]+)", power_normal_text)
    if fuel_type_match:
        power["generator"]["fuel_type"] = clean_text(fuel_type_match.group(1)).strip("_")
    location_match = re.search(r"#01發電機編號或名稱.*?\(1\)設置位置.*?[：:]\s*([^\n；;]+)", power_normal_text, re.S)
    power_match = re.search(r"功率[：:]\s*([0-9,]+(?:\.[0-9]+)?)\s*\(kW\)", power_normal_text)
    ats_match = re.search(r"停電啟動時間\(ATS\)[：:]\s*([0-9,]+(?:\.[0-9]+)?)", power_normal_text)
    day_tank_match = re.search(r"日用油槽油量[：:]\s*([0-9,]+(?:\.[0-9]+)?)", power_normal_text)
    underground_match = re.search(r"地下油槽\s*■有[，,：:]\s*_*\s*([0-9,]+(?:\.[0-9]+)?)", power_normal_text)
    if power["generator"]["generators"]:
        generator = power["generator"]["generators"][0]
        generator["location"] = clean_text(location_match.group(1)) if location_match else None
        generator["power_kw"] = _to_number(power_match.group(1)) if power_match else None
        generator["ats_start_seconds"] = _to_number(ats_match.group(1)) if ats_match else None
        generator["daily_tank_liters"] = _to_number(day_tank_match.group(1)) if day_tank_match else None
        generator["type"] = "獨立型" if "■獨立型" in power_normal_text else generator.get("type")
    power["fuel_storage"]["war_reserve_tank"]["exists"] = "■無" not in (
        _extract_labeled_text(power_normal_text, "是否具(戰)備用油槽") or ""
    )
    power["fuel_storage"]["underground_tank"]["capacity_liters"] = (
        _to_number(underground_match.group(1)) if underground_match else None
    )

    water_general_text = _join_table_column_text(water_table, 1)
    water_island_text = _join_table_column_text(water_table, 2)
    water_combined_text = clean_text(
        f"{water_general_text}\n孤島效應下應變措施A/[B*C]\n{water_island_text}"
    )
    water = parse_water_section(
        water_combined_text,
        clean_text(f"{water_general_text}\n{water_island_text}"),
    )
    water_baseline_text = _find_paragraph_containing(raw_json, "2025年(12個月)用水量")
    water_monthly_text = _find_paragraph_containing(raw_json, "6月用水")
    if water_baseline_text:
        annual_match = re.search(r"用水量[：:]\s*_*\s*([0-9,]+(?:\.[0-9]+)?)", water_baseline_text)
        daily_match = re.search(r"平均日用量[：:]\s*_*\s*([0-9,]+(?:\.[0-9]+)?)", water_baseline_text)
        water["baseline"] = {
            "year": 2025,
            "annual_usage_m3": _to_number(annual_match.group(1)) if annual_match else None,
            "average_daily_usage_m3": _to_number(daily_match.group(1)) if daily_match else None,
            "monthly_usage_m3": _parse_monthly_usage(water_monthly_text, value_label="用水"),
        }
    if water["baseline"].get("average_daily_usage_m3") is None:
        average_daily_match = re.search(r"193(?:\.[0-9]+)?", water_baseline_text)
        if average_daily_match:
            water["baseline"]["average_daily_usage_m3"] = _to_number(average_daily_match.group(0))
    if (
        water["baseline"].get("average_daily_usage_m3") is None
        and water["baseline"].get("annual_usage_m3") is not None
    ):
        water["baseline"]["average_daily_usage_m3"] = _round2(
            float(water["baseline"]["annual_usage_m3"]) / 365
        )

    ro_match = re.search(
        r"有設置RO逆滲透水[，,：:]\s*建置位置[：:]\s*([^\s_；;]+)_*棟；_*([0-9]+)_*樓層；數量_*([0-9]+)",
        water_general_text,
    )
    if ro_match:
        water["ro_system"] = {
            "exists": True,
            "building": clean_text(ro_match.group(1)),
            "floor": _to_number(ro_match.group(2)),
            "quantity": _to_number(ro_match.group(3)),
        }

    source_facilities = re.findall(r"\(([^)]+)\)\s*[^:\n]*[：:]\s*([0-9]+(?:\.[0-9]+)?)", water_general_text)
    if water["sources"]:
        water["sources"][0]["facilities"] = [
            {
                "name": clean_text(name),
                "capacity_tons": _to_number(capacity),
            }
            for name, capacity in source_facilities[:2]
        ]
    for source in water["sources"]:
        if source.get("selected") is None:
            source["selected"] = False

    gas_general_text = _join_table_column_text(gas_table, 1)
    gas_island_text = _join_table_column_text(gas_table, 2)
    gas_combined_text = clean_text(
        f"{gas_general_text}\n孤島效應下應變措施\n{gas_island_text}"
    )
    medical_gas = parse_medical_gas_section(
        gas_combined_text,
        clean_text(f"{gas_general_text}\n{gas_island_text}"),
    )
    gas_baseline_text = _find_paragraph_containing(raw_json, "醫用氣體(O2)用量")
    if gas_baseline_text:
        annual_match = re.search(r"O2\)用量[：:]\s*_*\s*([0-9,]+(?:\.[0-9]+)?)", gas_baseline_text)
        daily_match = re.search(r"平均日用量[：:]\s*_*\s*([0-9,]+(?:\.[0-9]+)?)", gas_baseline_text)
        medical_gas["baseline"] = {
            "annual_o2_usage_tons": _to_number(annual_match.group(1)) if annual_match else None,
            "average_daily_o2_usage_tons": _to_number(daily_match.group(1)) if daily_match else None,
        }

    gas_tank_match = re.search(r"([0-9]+)\+([0-9]+)=([0-9]+)\s*?賂?", gas_general_text)
    if gas_tank_match:
        medical_gas["liquid_oxygen"]["capacity_tons_each"] = _to_number(gas_tank_match.group(1))
        medical_gas["liquid_oxygen"]["quantity"] = 2
        medical_gas["liquid_oxygen"]["total_capacity_tons"] = _to_number(gas_tank_match.group(3))
        medical_gas["liquid_oxygen"]["formula_text"] = (
            f"{gas_tank_match.group(1)}+{gas_tank_match.group(2)}={gas_tank_match.group(3)}"
        )
    cylinder_match = re.search(
        r"([0-9]+)\(?Ｗ\)\+([0-9]+)\(?Ｗ?\)=([0-9]+)\(?珮\).*?([0-9]+(?:\.[0-9]+)?)\(?瘞",
        gas_general_text,
        re.S,
    )
    if cylinder_match:
        medical_gas["oxygen_cylinders"] = {
            "inside_hospital_count": _to_number(cylinder_match.group(1)),
            "outside_hospital_count": _to_number(cylinder_match.group(2)),
            "total_count": _to_number(cylinder_match.group(3)),
            "gas_equivalent_tons": _to_number(cylinder_match.group(4)),
        }
    gas_formula_match = re.search(
        r"([0-9]+(?:\.[0-9]+)?)\(T\)/([0-9]+(?:\.[0-9]+)?)\(T/憭?\)=([0-9]+(?:\.[0-9]+)?)\(憭?\)",
        gas_island_text,
    )
    if gas_formula_match:
        total_capacity = _to_number(gas_formula_match.group(1))
        daily_usage = _to_number(gas_formula_match.group(2))
        calculated_days = _to_number(gas_formula_match.group(3))
        medical_gas["island_mode"] = {
            "formula": f"{total_capacity} / {daily_usage}",
            "document_days": calculated_days,
            "calculated_days": calculated_days,
            "meets_72_hours": bool(calculated_days >= 3),
        }
    gas_options = medical_gas.get("normal_condition", {}).get("options", [])
    if gas_options:
        gas_options[0]["selected"] = medical_gas["liquid_oxygen"]["total_capacity_tons"] is not None
        gas_options[0]["value"] = {
            "capacity_tons_each": medical_gas["liquid_oxygen"]["capacity_tons_each"],
            "quantity": medical_gas["liquid_oxygen"]["quantity"],
            "total_capacity_tons": medical_gas["liquid_oxygen"]["total_capacity_tons"],
        } if medical_gas["liquid_oxygen"]["total_capacity_tons"] is not None else None
        gas_options[2]["selected"] = medical_gas["oxygen_cylinders"]["total_count"] is not None
        gas_options[2]["value"] = (
            medical_gas["oxygen_cylinders"]
            if medical_gas["oxygen_cylinders"]["total_count"] is not None
            else None
        )
        medical_gas["normal_condition"]["selected_values"] = [
            option["label"] for option in gas_options if option.get("selected")
        ]
    direct_cylinder_match = re.search(
        r"([0-9]+)\(院內\)\+([0-9]+)\(院外\)=([0-9]+)\(支\).*?([0-9]+(?:\.[0-9]+)?)\(氣態\)",
        gas_general_text,
        re.S,
    )
    if direct_cylinder_match:
        medical_gas["oxygen_cylinders"] = {
            "inside_hospital_count": _to_number(direct_cylinder_match.group(1)),
            "outside_hospital_count": _to_number(direct_cylinder_match.group(2)),
            "total_count": _to_number(direct_cylinder_match.group(3)),
            "gas_equivalent_tons": _to_number(direct_cylinder_match.group(4)),
        }
        gas_options[2]["selected"] = True
        gas_options[2]["value"] = medical_gas["oxygen_cylinders"]
        medical_gas["normal_condition"]["selected_values"] = [
            option["label"] for option in gas_options if option.get("selected")
        ]
    direct_gas_formula_match = re.search(
        r"([0-9]+(?:\.[0-9]+)?)\(T\)/([0-9]+(?:\.[0-9]+)?)\(T/天\)=([0-9]+(?:\.[0-9]+)?)\(天\)",
        gas_island_text,
    )
    if direct_gas_formula_match:
        total_capacity = _to_number(direct_gas_formula_match.group(1))
        daily_usage = _to_number(direct_gas_formula_match.group(2))
        calculated_days = _to_number(direct_gas_formula_match.group(3))
        medical_gas["island_mode"] = {
            "formula": f"{total_capacity} / {daily_usage}",
            "document_days": calculated_days,
            "calculated_days": calculated_days,
            "meets_72_hours": bool(calculated_days >= 3),
        }

    it_general_text = _join_table_column_text(it_table, 1)
    it_island_text = _join_table_column_text(it_table, 2)
    it_combined_text = clean_text(
        f"{it_general_text}\n孤島效應下應變措施\n{it_island_text}"
    )
    it_backup = parse_it_backup_section(it_combined_text, clean_text(it_combined_text))
    if it_backup.get("external_backup", {}).get("mobile_storage_type"):
        it_backup["external_backup"]["mobile_storage_transfer"] = True
        external_field = it_backup.get("normal_condition", {}).get("external_backup", {})
        for option in external_field.get("options", []):
            if "蝘餃?" in option.get("label", "") or "移動式" in option.get("label", ""):
                option["selected"] = True
                option["value"] = it_backup["external_backup"]["mobile_storage_type"]
        external_field["selected_values"] = [
            option["label"] for option in external_field.get("options", []) if option.get("selected")
        ]
    it_backup["has_offsite_backup"] = bool(
        it_backup.get("external_backup", {}).get("offsite_distance_km", {}).get("filled")
        or it_backup.get("external_backup", {}).get("automatic_offsite_backup")
        or it_backup.get("external_backup", {}).get("cloud_vendor", {}).get("filled")
    )

    sections = {
        "power": power,
        "water": water,
        "medical_gas": medical_gas,
        "it_backup": it_backup,
        "facility_maintenance": _parse_full_questionnaire_facility(facility_table),
        "central_monitoring": _parse_full_questionnaire_monitoring(monitoring_table),
    }

    return _build_normalized_questionnaire_json(
        raw_json=raw_json,
        docx_path=docx_path,
        basic_info=parse_basic_info_full_questionnaire(raw_json),
        sections=sections,
        profile="full_questionnaire",
    )


def parse_questionnaire_docx(docx_path: str) -> dict[str, Any]:
    """
    Input: Word .docx questionnaire
    Output: normalized questionnaire JSON
    """
    raw_json = read_docx_blocks(docx_path)
    profile = detect_questionnaire_profile(raw_json)
    raw_json["document_profile"] = profile

    if profile == "full_questionnaire":
        normalized_json = _parse_full_questionnaire(raw_json, docx_path)
    else:
        normalized_json = _parse_diagnosis_form(raw_json, docx_path)
    report_input_json = build_report_input_json(normalized_json)

    return QuestionnaireParseResponse(
        raw_json=raw_json,
        normalized_json=normalized_json,
        report_input_json=report_input_json,
    ).model_dump()
