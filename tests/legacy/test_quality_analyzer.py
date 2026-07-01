"""
Unit tests for business_logic/quality_analyzer.py
Mocks ollama module to avoid network calls.
Tests score parsing, edge cases, and None confidence handling.
"""

import sys
import os
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

# Defensive guard for COM mocks
if "win32com" not in sys.modules:
    _wc = MagicMock()
    _wcc = MagicMock()
    _wc.client = _wcc
    sys.modules.update({"win32com": _wc, "win32com.client": _wcc, "pythoncom": MagicMock()})

from business_logic.quality_analyzer import QualityAnalyzer
from domain.models import DocumentContent, QualityAnalysisResult
from domain.enums import QualityLevel


def _make_analyzer(response_text_1="", response_text_2=""):
    """Return QualityAnalyzer with mocked ollama module."""
    with patch("ollama.Client") as mock_client_cls:
        mock_client = MagicMock()
        mock_client_cls.return_value = mock_client

        analyzer = QualityAnalyzer()

        # Mock the module-level ollama.generate calls
        mock_ollama = MagicMock()
        mock_ollama.generate.side_effect = [
            {"response": response_text_1},
            {"response": response_text_2},
        ]
        analyzer.ollama = mock_ollama

    return analyzer


def _doc(paragraphs=None):
    return DocumentContent(
        word_count=1000,
        char_count=5000,
        title="Test document",
        paragraphs=paragraphs
        or [
            "Introducción al estudio.",
            "Metodología aplicada.",
            "Resultados obtenidos.",
            "Conclusiones del trabajo.",
        ],
    )


GOOD_RESPONSE_1 = """
**1. Claridad del argumento** [Puntuación: 8/10]
El argumento central es claro y bien estructurado.

**2. Coherencia** [Puntuación: 7/10]
Las ideas se conectan de manera lógica.
"""

GOOD_RESPONSE_2 = """
**1. Argumentación** [Puntuación: 8/10]
Los argumentos están bien fundamentados con evidencia.

**2. Conclusiones** [Puntuación: 7/10]
Las conclusiones son pertinentes y derivadas del análisis.
"""


class TestQualityAnalyzerHappyPath(unittest.TestCase):
    def test_returns_quality_analysis_result(self):
        analyzer = _make_analyzer(GOOD_RESPONSE_1, GOOD_RESPONSE_2)
        result = analyzer.analyze_quality(_doc(), None)
        self.assertIsInstance(result, QualityAnalysisResult)

    def test_overall_score_in_range(self):
        analyzer = _make_analyzer(GOOD_RESPONSE_1, GOOD_RESPONSE_2)
        result = analyzer.analyze_quality(_doc(), None)
        self.assertGreaterEqual(result.overall_score, 0.0)
        self.assertLessEqual(result.overall_score, 10.0)

    def test_dimension_scores_keys_present(self):
        analyzer = _make_analyzer(GOOD_RESPONSE_1, GOOD_RESPONSE_2)
        result = analyzer.analyze_quality(_doc(), None)
        for key in ["claridad", "coherencia", "argumentacion", "conclusiones"]:
            self.assertIn(key, result.dimension_scores)

    def test_quality_level_is_enum(self):
        analyzer = _make_analyzer(GOOD_RESPONSE_1, GOOD_RESPONSE_2)
        result = analyzer.analyze_quality(_doc(), None)
        self.assertIsInstance(result.quality_level, QualityLevel)


class TestQualityAnalyzerEdgeCases(unittest.TestCase):
    def test_empty_response_uses_defaults(self):
        """Empty LLM response should not crash — defaults to 7.0."""
        analyzer = _make_analyzer("", "")
        result = analyzer.analyze_quality(_doc(), None)
        self.assertIsNotNone(result)
        self.assertEqual(result.overall_score, 7.0)

    def test_malformed_response_no_scores(self):
        """Response without score format falls back to narrative inference."""
        analyzer = _make_analyzer("El texto es bueno y adecuado.", "Es suficiente y aceptable.")
        result = analyzer.analyze_quality(_doc(), None)
        self.assertIsNotNone(result)
        self.assertGreater(result.overall_score, 0.0)

    def test_ollama_exception_returns_default(self):
        """LLM exception returns default 7.0 result."""
        with patch("ollama.Client"):
            analyzer = QualityAnalyzer()
            mock_ollama = MagicMock()
            mock_ollama.generate.side_effect = Exception("LLM unavailable")
            analyzer.ollama = mock_ollama

        result = analyzer.analyze_quality(_doc(), None)
        self.assertIsNotNone(result)
        self.assertEqual(result.overall_score, 7.0)

    def test_short_document_uses_full_text(self):
        """Documents < 400 words should use full text (no sampling crash)."""
        analyzer = _make_analyzer(GOOD_RESPONSE_1, GOOD_RESPONSE_2)
        short_doc = _doc(["Texto corto."])
        result = analyzer.analyze_quality(short_doc, None)
        self.assertIsNotNone(result)


class TestParseScoreExtraction(unittest.TestCase):
    """Test _parse_llm_response directly."""

    def setUp(self):
        with patch("ollama.Client"):
            self.analyzer = QualityAnalyzer()

    def test_score_extraction_from_bracket_format(self):
        text = "**1. Claridad del argumento** [Puntuación: 9/10]\nTexto de análisis."
        result = self.analyzer._parse_llm_response(text)
        self.assertEqual(result["claridad"]["score"], 9.0)

    def test_score_extraction_inline_format(self):
        text = "**Coherencia** 8/10\nTexto de análisis de coherencia."
        result = self.analyzer._parse_llm_response(text)
        self.assertEqual(result["coherencia"]["score"], 8.0)

    def test_score_clamped_to_10(self):
        text = "**Claridad del argumento** [Puntuación: 15/10]\nExcelente."
        result = self.analyzer._parse_llm_response(text)
        self.assertLessEqual(result["claridad"]["score"], 10.0)

    def test_narrative_inference_excelente(self):
        """'excelente' without score format → infers 8.5."""
        text = "**Argumentación**\nEl argumento es excelente y muy bien fundamentado."
        result = self.analyzer._parse_llm_response(text)
        self.assertEqual(result["argumentacion"]["score"], 8.5)

    def test_narrative_inference_deficiente(self):
        """'deficiente' without score format → infers 4.0."""
        text = "**Conclusiones**\nLas conclusiones son deficientes y débiles."
        result = self.analyzer._parse_llm_response(text)
        self.assertEqual(result["conclusiones"]["score"], 4.0)

    def test_empty_text_returns_defaults(self):
        result = self.analyzer._parse_llm_response("")
        for key in ["claridad", "coherencia", "argumentacion", "conclusiones"]:
            self.assertEqual(result[key]["score"], 7.0)
            self.assertEqual(result[key]["feedback"], "No disponible")


if __name__ == "__main__":
    unittest.main()
