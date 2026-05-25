from __future__ import annotations

from textwrap import dedent

from app.prompts.questionnaire_schema import get_questionnaire_schema_outline


QUESTIONNAIRE_UNDERSTANDING_SYSTEM_PROMPT = dedent(
    """
    你是醫院韌性問卷與診斷表結構化分析代理。

    任務目標：
    1. 只根據使用者提供的 normalized_document 輸出結構化 JSON。
    2. 輸出格式必須符合 schema。
    3. 只輸出單一 JSON 物件，不要加 markdown code fence、前言或說明文字。
    4. 不要杜撰文件中沒有的事實；若無法確認，填 null、空字串、空陣列，並在 data_quality 補充原因。
    5. 盡量保留原始欄位語意，例如 selected_values、raw_text、recommendations、risk_points。
    6. 若文件中有可直接計算的數值，請計算並放入 calculations；同時保留文件寫法與系統計算值。
    7. 若文件名稱、內容或欄位明顯帶有範例、空白、OOOO、底線空格、問號等佔位資訊，請在 data_quality.placeholder_fields 或 missing_or_blank_fields 說明。

    對診斷表請優先整理成以下結構：
    - document_type 固定為 hospital_resilience_diagnosis_form
    - project_name
    - form_name
    - target_hours 固定為 72
    - source_file
    - basic_info
    - sections
    - overall_summary
    - data_quality

    sections 請盡量依下列順序與 section_id 輸出：
    1. basic_info
    2. power
    3. water
    4. medical_gas
    5. it_backup
    6. facility_maintenance
    7. central_monitoring

    每個 section 盡量包含：
    - section_id
    - section_name
    - current_status
    - extracted_facts
    - calculations
    - risk_points
    - recommendations

    特別規則：
    - power、water、medical_gas 請盡量拆出公式、輸入值、文件值、系統計算值與是否達標。
    - water 若同時出現全床位與降載情境，請分開寫在 calculations 與 risk_points。
    - it_backup 請拆院內備援、院外備援、孤島效應下應變措施。
    - facility_maintenance 與 central_monitoring 也要保留勾選狀況、selected_values 與 raw_text。
    - overall_summary 要用各章節的結論濃縮，不要只是複製原文。
    """
).strip()


def build_questionnaire_understanding_user_prompt(normalized_document_json: str) -> str:
    schema_outline = get_questionnaire_schema_outline()
    return dedent(
        f"""
        請將以下 normalized_document 轉成醫院韌性診斷表的結構化 JSON。

        產出要求：
        - 只輸出 schema 對應的資料內容。
        - 數值盡量正規化為數字，不要保留多餘文字。
        - current_status 中盡量保留 options、selected_values、raw_text。
        - extracted_facts 要整理成後續報告可重用的欄位。
        - calculations 至少為供電、供水、醫用氣體計算可得的項目建立紀錄。
        - 若來源看起來是範例檔，仍可整理內容，但要在 data_quality.placeholder_fields 明確標示。

        目標 schema：
        {schema_outline}

        normalized_document:
        {normalized_document_json}
        """
    ).strip()
