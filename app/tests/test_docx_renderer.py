from __future__ import annotations

import unittest
from pathlib import Path

from app.models.schemas import ReportJSON
from app.services.docx_renderer import DocxRenderer


ROOT = Path(__file__).resolve().parents[2]


class DocxRendererDiagnosisTests(unittest.TestCase):
    def test_uses_single_llm_diagnosis_paragraph_instead_of_fallback(self) -> None:
        renderer = DocxRenderer(output_dir=ROOT / "outputs")
        paragraph = "依據問卷資料，該院供電、供水與供氧均已具基礎備援能力，惟仍需確認重大災害下切換與監控機制。"
        report = ReportJSON(
            title="test",
            summary="test",
            background="test",
            diagnosis_paragraphs=[
                {
                    "text": paragraph,
                    "evidence": [
                        {
                            "source_name": "問卷.docx",
                            "fact_type": "paragraph",
                            "content": "供電、供水與供氧均已具基礎備援能力",
                        }
                    ],
                }
            ],
        )

        self.assertEqual(
            renderer._build_diagnosis_paragraphs(report, {}),
            [paragraph],
        )

    def test_missing_llm_diagnosis_does_not_use_template_observation(self) -> None:
        renderer = DocxRenderer(output_dir=ROOT / "outputs")
        report = ReportJSON(title="test", summary="test", background="test")

        paragraphs = renderer._build_diagnosis_paragraphs(report, {})

        self.assertEqual(len(paragraphs), 1)
        self.assertIn("依現有資料尚無法完整判定", paragraphs[0])
        self.assertNotIn("各運作量能均可能達72小時", paragraphs[0])

    def test_drops_llm_diagnosis_without_evidence(self) -> None:
        renderer = DocxRenderer(output_dir=ROOT / "outputs")
        report = ReportJSON(
            title="test",
            summary="test",
            background="test",
            diagnosis_paragraphs=["沒有來源的診斷段落"],
        )

        paragraphs = renderer._build_diagnosis_paragraphs(report, {})

        self.assertEqual(len(paragraphs), 1)
        self.assertIn("依現有資料尚無法完整判定", paragraphs[0])


if __name__ == "__main__":
    unittest.main()
