from __future__ import annotations

import unittest
from pathlib import Path

from app.parsers.questionnaire_docx_parser import (
    detect_questionnaire_profile,
    detect_section,
    parse_checkbox_options,
    parse_questionnaire_docx,
    read_docx_blocks,
)


ROOT = Path(__file__).resolve().parents[2]
DIAGNOSIS_DOC = ROOT / "B_醫院專家診斷表(範例)v0.docx"
FULL_QUESTIONNAIRE_DOC = ROOT / "5566.docx"


class QuestionnaireDocxParserTests(unittest.TestCase):
    def test_detect_section(self) -> None:
        self.assertEqual(detect_section("供電相關"), "power")
        self.assertEqual(detect_section("供水相關"), "water")
        self.assertEqual(detect_section("醫用氣體"), "medical_gas")
        self.assertEqual(detect_section("資訊系統備援"), "it_backup")

    def test_parse_checkbox_options(self) -> None:
        parsed = parse_checkbox_options("■供電\n□供水\n□供氣\n■IT機房", question="test")
        self.assertEqual(parsed["selected_values"], ["供電", "IT機房"])
        self.assertEqual(len(parsed["options"]), 4)
        self.assertTrue(parsed["options"][0]["selected"])
        self.assertFalse(parsed["options"][1]["selected"])

    def test_read_docx_blocks(self) -> None:
        raw = read_docx_blocks(str(DIAGNOSIS_DOC))
        self.assertEqual(raw["document_type"], "hospital_resilience_questionnaire")
        self.assertGreaterEqual(len(raw["paragraphs"]), 2)
        self.assertEqual(len(raw["tables"]), 1)

    def test_detect_profiles(self) -> None:
        diagnosis_raw = read_docx_blocks(str(DIAGNOSIS_DOC))
        full_raw = read_docx_blocks(str(FULL_QUESTIONNAIRE_DOC))

        self.assertEqual(detect_questionnaire_profile(diagnosis_raw), "diagnosis_form")
        self.assertEqual(detect_questionnaire_profile(full_raw), "full_questionnaire")

    def test_parse_diagnosis_form(self) -> None:
        result = parse_questionnaire_docx(str(DIAGNOSIS_DOC))
        normalized = result["normalized_json"]

        self.assertEqual(result["raw_json"]["source_file"], "B_醫院專家診斷表(範例)v0.docx")
        self.assertEqual(normalized["document_profile"], "diagnosis_form")
        self.assertEqual(normalized["basic_info"]["hospital_name"], "OOOO醫院")
        self.assertEqual(normalized["sections"]["power"]["generator"]["generator_count"], 1)
        self.assertAlmostEqual(
            normalized["sections"]["water"]["island_mode"]["reduced_bed_scenario"]["calculated_days"],
            9.72,
            places=1,
        )
        self.assertEqual(
            normalized["sections"]["water"]["normal_condition"]["options"][5]["label"],
            "每日每床用水量",
        )
        self.assertEqual(
            normalized["sections"]["water"]["normal_condition"]["options"][5]["value"],
            1000,
        )
        self.assertEqual(
            normalized["sections"]["water"]["island_mode_response"]["options"][0]["label"],
            "孤島效應下最大供水可達天數",
        )
        self.assertEqual(
            normalized["sections"]["water"]["island_mode_response"]["options"][0]["value"],
            2.84,
        )
        self.assertTrue(normalized["sections"]["water"]["sources"][0]["selected"])
        self.assertFalse(normalized["sections"]["water"]["sources"][1]["selected"])
        self.assertTrue(normalized["sections"]["medical_gas"]["island_mode"]["meets_72_hours"])
        self.assertEqual(
            normalized["sections"]["it_backup"]["external_backup"]["mobile_storage_type"],
            "磁帶",
        )
        self.assertFalse(normalized["sections"]["it_backup"]["has_offsite_backup"])
        self.assertIsNone(
            normalized["sections"]["it_backup"]["external_backup"]["offsite_distance_km"]["value"]
        )
        self.assertFalse(
            normalized["sections"]["it_backup"]["external_backup"]["backup_type"]["filled"]
        )
        self.assertIn("overall_summary", normalized)
        self.assertTrue(
            any("OOOO醫院" in warning for warning in normalized["data_quality"]["warnings"])
        )

    def test_parse_full_questionnaire(self) -> None:
        result = parse_questionnaire_docx(str(FULL_QUESTIONNAIRE_DOC))
        normalized = result["normalized_json"]

        self.assertEqual(normalized["document_profile"], "full_questionnaire")
        self.assertEqual(normalized["basic_info"]["hospital_name"], "OOOO醫院")
        self.assertEqual(normalized["basic_info"]["diagnosis_date"], "2026.02.06")
        self.assertEqual(normalized["sections"]["power"]["baseline"]["annual_usage_kwh"], 6672042)
        self.assertEqual(normalized["sections"]["power"]["generator"]["generator_count"], 1)
        self.assertEqual(
            normalized["sections"]["water"]["baseline"]["annual_usage_m3"],
            70762,
        )
        self.assertEqual(
            normalized["sections"]["water"]["sources"][0]["type"],
            "自來水機構供水",
        )
        self.assertTrue(normalized["sections"]["water"]["sources"][0]["selected"])
        self.assertFalse(normalized["sections"]["water"]["sources"][1]["selected"])
        self.assertEqual(
            normalized["sections"]["water"]["ro_system"]["building"],
            "綜合醫療大樓",
        )
        self.assertEqual(
            normalized["sections"]["medical_gas"]["liquid_oxygen"]["total_capacity_tons"],
            40,
        )
        self.assertEqual(
            normalized["sections"]["it_backup"]["external_backup"]["mobile_storage_type"],
            "磁帶",
        )
        self.assertTrue(
            normalized["sections"]["it_backup"]["external_backup"]["mobile_storage_transfer"]
        )
        self.assertFalse(normalized["sections"]["it_backup"]["has_offsite_backup"])
        self.assertIn("供電", normalized["sections"]["facility_maintenance"]["selected_values"])


    def test_report_input_json_shape(self) -> None:
        diagnosis_doc = next(ROOT.glob("B_*.docx"))
        diagnosis_result = parse_questionnaire_docx(str(diagnosis_doc))
        diagnosis_report = diagnosis_result["report_input_json"]

        self.assertEqual(diagnosis_report["document_type"], "hospital_resilience_diagnosis_form")
        self.assertEqual(
            [section["section_id"] for section in diagnosis_report["sections"]],
            [
                "basic_info",
                "power",
                "water",
                "medical_gas",
                "it_backup",
                "facility_maintenance",
                "central_monitoring",
            ],
        )
        self.assertEqual(
            diagnosis_report["sections"][1]["extracted_facts"]["generator_count"],
            1,
        )
        self.assertIn("missing_or_blank_fields", diagnosis_report["data_quality"])

        full_result = parse_questionnaire_docx(str(FULL_QUESTIONNAIRE_DOC))
        full_report = full_result["report_input_json"]
        self.assertEqual(full_report["document_type"], "hospital_resilience_questionnaire")
        self.assertEqual(full_report["document_profile"], "full_questionnaire")


if __name__ == "__main__":
    unittest.main()
