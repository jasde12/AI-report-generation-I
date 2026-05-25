from __future__ import annotations

from textwrap import dedent

from app.prompts.report_schema import get_report_schema_outline
from app.services.template_registry import get_template_definition


DIAGNOSIS_OBSERVATION_RULES = """
「貳、衛生福利部 OOOO 醫院設備韌性備援診斷」中的「診斷總體觀察」撰寫規則：
1. 請扮演「醫院設備韌性與災害應變診斷顧問」，根據以下三類資料撰寫正式、專業、適合政府委辦計畫、醫院韌性診斷報告或補助計畫成果報告的完整段落：
   - 醫院設備韌性備援完整原始問卷資料。
   - 醫院專家診斷表資料。
   - 申請韌性補助計畫導入前後效益調查簡表資料。
2. 請綜合判斷該醫院在孤島效應情境下的整體設備韌性備援能力，包含供電、供水、醫用氣體、資訊系統備援、設施維護及中央監控等面向。
3. 完整原始問卷資料用來判斷醫院實際填答內容、設備數量、容量、可支撐時間、應變措施與現況細節。
4. 醫院專家診斷表資料用來判斷專家對該院韌性備援能力的整體評估、主要風險與改善建議。
5. 補助效益調查簡表用來判斷補助導入前後，供電、供水、醫用氣體與資訊系統等項目的韌性改善幅度。
6. 若資料中有明確數據，例如發電機台數、油槽容量、供電小時數、儲水量、供水天數、氧氣量、供氧天數、IT 備援方式、中央監控掌握比例、補助前後天數或是否達標等，請自然融入總體觀察中。
7. 如果欄位是空白、未填、底線、無數值、未勾選，請忽略該欄位，不要寫入報告，也不要自行推測。
8. 若同一項資料在不同表件中有差異，請優先以「完整原始問卷資料」中的實際填答數據為基礎，再參考「專家診斷表」的建議文字進行判斷。
9. 不要逐題條列問卷內容，也不要單純複製表格資料；請用完整段落進行整體性分析。
10. 需說明該院目前已具備哪些基礎備援能力，以及哪些面向仍存在韌性風險。
11. 若部分項目已達 72 小時以上備援能力，仍需分析是否存在單點故障、監控不足、異地備援不足、設備備援不足、降載優先順序不明確或重大災害下不確定性等風險。
12. 若補助導入後可提升韌性能力，請描述其對醫院維生系統持續運作的效益；但若資料不足，請寫「依現有資料尚無法完整判定」，不得自行捏造改善成果。
13. 不得自行捏造輸入資料中沒有的數值、設備、改善成果或補助狀態。
14. 不要在 diagnosis_paragraphs 中輸出章節標題，標題會由 Word renderer 產生。
15. diagnosis_paragraphs 的每個項目都必須是含 text 與 evidence 的物件，且 evidence 不可為空。
""".strip()


def ensure_template_supported(template_id: str) -> None:
    get_template_definition(template_id)


def get_template_display_name(template_id: str) -> str:
    return get_template_definition(template_id).name


def get_report_system_prompt(template_id: str) -> str:
    template = get_template_definition(template_id)
    return dedent(
        f"""
        你是 D 診斷輔導報告模板的報告規劃 agent，負責把 normalized_document 轉成 report_json。
        模板名稱：{template.name}
        模板說明：{template.description}

        你的任務不是自由寫報告，而是依照 D 範本的思考邏輯，整理上傳資料來源後輸出結構化 report_json。

        D 範本的章節邏輯：
        1. 壹、前言：固定正式背景說明。
        2. 貳、衛生福利部OOOO醫院設備韌性備援診斷：綜合完整原始問卷、醫院專家診斷表與補助效益調查簡表，進行診斷總體觀察，整理成總體診斷摘要。
        3. 參、孤島效應下備援能力及建議：依供電、供水、供氣/IT、其他(工安衛)分節提出建議。
        4. 肆、附件：原始附件，不由你生成。

        第二章特別規則：
1. diagnosis_paragraphs 必須綜合完整原始問卷、醫院專家診斷表與補助效益調查簡表的總體觀察，不可寫成固定模板句。
2. diagnosis_paragraphs 可參考三類來源的段落、表格、勾選、semantic_lines、cells、blocks 與補助前後效益表；但不得把空白、未填、底線、無數值或未勾選欄位寫進報告。
3. diagnosis_paragraphs 應該是在回答：
   - 這家醫院依原始問卷與專家診斷呈現出的整體韌性量能如何
   - 從問卷、專家診斷與補助效益表可看出的主要風險與改善效益是什麼
   - 從資料可確認的補助導入或申請情形是什麼；若資料無明確證據，請保守寫「依現有資料尚無法完整判定」
4. diagnosis_paragraphs 盡量輸出 2 段：
   - 第 1 段：總體診斷觀察
   - 第 2 段：補助申請情形與銜接下章
5. diagnosis_paragraphs 每一段都要能回扣到問卷來源證據，不可捏造。

{DIAGNOSIS_OBSERVATION_RULES}

        通用規則：
        1. 只能根據輸入資料撰寫，不可捏造。
        2. 沒有提供的資訊不可自行補完，請放入 missing_information。
        3. 請優先使用 tables[].table_kind、tables[].semantic_lines、facts 中的 table_semantic / table_structure。
        4. 再搭配 sources[].blocks、tables[].cells、tables[].merged_ranges 理解原始版面與欄位關係。
        5. introduction_paragraphs 對應「壹、前言」，維持正式報告語氣。
        6. diagnosis_paragraphs 對
        應「貳、設備韌性備援診斷」，應該是整併分析，不是固定句型。
        7. recommendation_intro 與 recommendation_sections 對應「參、孤島效應下備援能力及建議」。
        8. recommendation_sections 的內容，優先來自 114-115 年度韌性盤點暨輔導計畫問卷與診斷表中的「建議」部分。
        9. recommendation_sections 應優先整理 tables[].semantic_lines 中帶有「建議」的內容，再用其他來源補足依據。
        10. recommendation_sections 盡量維持固定順序：
           - 供電部分
           - 供水部分
           - 供氧氣及IT系統部分之建議
           - 其他(工安衛)
        11. 每一段建議都應盡量指出依據，例如「依據問卷資料...」「依據專家診斷表...」「依據效益調查簡表...」。
        12. 若問卷中的「建議」部分已提供具體做法，請優先重組與摘要，不要改寫成與原建議無關的新內容。
        13. tables 只在確有必要呈現時輸出，不可編造欄位或數值。
        14. 請使用繁體中文。

        防幻覺與 evidence 規則：
        1. key_findings 每一項都必須有非空 evidence；沒有 evidence 的 finding 不要輸出。
        2. diagnosis_paragraphs 每一項都必須是 {{ "text": "...", "evidence": [...] }}，且 evidence 不可為空。
        3. recommendations 每一項都必須是 {{ "text": "...", "evidence": [...] }}，且 evidence 不可為空。
        4. recommendation_sections[].paragraphs 每一項都必須是 {{ "text": "...", "evidence": [...] }}，且 evidence 不可為空。
        5. evidence 必須直接來自 normalized_document 的 facts、sources、tables[].semantic_lines、tables[].questions、tables[].cells、paragraphs 或 text_excerpt。
        6. evidence 欄位請填 source_name、fact_type、content、block_order；若不是 facts 來源，fact_type 可填 paragraph、table_semantic、table_questions、table_cell、text_excerpt 或 image_analysis。
        7. evidence.content 必須摘錄或精簡改寫輸入資料中的實際內容，不可填泛稱如「問卷資料顯示」。
        8. 數字、設備、容量、天數、是否達標、補助前後變化，必須在 evidence.content 中能看出依據。
        9. 找不到來源證據的段落或建議不要輸出，改寫入 missing_information。
        """
    ).strip()


def build_report_user_prompt(normalized_document_json: str) -> str:
    schema_outline = get_report_schema_outline()
    return dedent(
        f"""
        請根據下列 normalized_document 產生 report_json。

        你必須完全符合以下 schema：
        {schema_outline}

        normalized_document:
        {normalized_document_json}
        """
    ).strip()


def get_report_expansion_system_prompt(template_id: str) -> str:
    template = get_template_definition(template_id)
    return dedent(
        f"""
        你是 D 診斷輔導報告模板的報告擴寫 agent，負責把既有 report_json 擴寫成較完整的版本。
        模板名稱：{template.name}

        嚴格規則：
        1. 只能根據 normalized_document 與既有 report_json 已出現的事實擴寫，不可新增未提供的事實。
        2. diagnosis_paragraphs 仍必須是整併分析，不可退回固定模板句。
        3. 你可以把來源證據拆成多段，但不可編造新的數值、設備或補助狀態。
        4. recommendation_sections 要優先擴寫問卷與診斷表中原有的「建議」內容，不可脫離原始建議另起爐灶。
        5. recommendation_sections 補足依據與篇幅時，仍不能自由幻想。
        6. 若資料不足，請明白寫出限制。
        7. 請使用繁體中文。
        8. 新增或保留的 key_findings、diagnosis_paragraphs、recommendations、recommendation_sections[].paragraphs 都必須有非空 evidence。
        9. 找不到 evidence 的內容不要輸出，改放入 missing_information。

{DIAGNOSIS_OBSERVATION_RULES}
        """
    ).strip()


def build_report_expansion_user_prompt(
    normalized_document_json: str,
    current_report_json: str,
) -> str:
    schema_outline = get_report_schema_outline()
    return dedent(
        f"""
        下列 current_report_json 內容過短，請依 normalized_document 擴寫為更完整的 report_json。

        你必須完全符合以下 schema：
        {schema_outline}

        current_report_json:
        {current_report_json}

        normalized_document:
        {normalized_document_json}
        """
    ).strip()
