"""
Unit tests for business_logic/article_classifier.py
Mocks ollama.Client to avoid network calls.
Tests the 19-case classification rule and signal helpers.
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

from business_logic.article_classifier import ArticleClassifier
from domain.models import DocumentContent, Reference
from domain.enums import ArticleType


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_classifier(generate_response="S4: SI\nS5: SI\nS6: SI"):
    """Return ArticleClassifier with mocked Ollama client."""
    with patch("ollama.Client") as mock_client_cls:
        mock_client = MagicMock()
        mock_client.generate.return_value = {"response": generate_response}
        mock_client_cls.return_value = mock_client
        classifier = ArticleClassifier()
        classifier.client = mock_client
    return classifier


def _scientific_doc(references=20, recent_years=True, methodological=True):
    """DocumentContent that should trigger most scientific signals."""
    year_suffix = "2022" if recent_years else "1990"
    refs = [
        Reference(text=f"Autor, A. ({year_suffix}). Título {i}. Revista, {i + 1}, 1-10.")
        for i in range(references)
    ]
    method_paras = (
        [
            "Este estudio analiza los efectos de la integración utilizando metodología cuantitativa.",
            "La hipótesis central sostiene que la cohesión aumenta con formación conjunta.",
            "Se aplicó encuesta a muestra de 240 sujetos con diseño cuasi-experimental.",
            "El análisis estadístico muestra correlación significativa (r=0.72, p<0.01).",
            "Los resultados demuestran validación de la hipótesis mediante simulación.",
        ]
        if methodological
        else ["Texto de opinión sin metodología."]
    )

    return DocumentContent(
        word_count=5000,
        char_count=30000,
        title="Análisis cuantitativo de la cohesión institucional",
        paragraphs=method_paras,
        references=refs,
    )


# ---------------------------------------------------------------------------
# Signal unit tests (deterministic signals)
# ---------------------------------------------------------------------------


class TestSignalReferenceCount(unittest.TestCase):
    def setUp(self):
        self.classifier = _make_classifier()

    def test_signal_s2a_true_when_references_gte_12(self):
        refs = [Reference(text=f"Ref {i} (2020).") for i in range(15)]
        doc_content = DocumentContent(
            word_count=1000, char_count=5000, references=refs, paragraphs=["x"]
        )
        self.assertTrue(self.classifier._signal_reference_count(doc_content))

    def test_signal_s2a_false_when_references_lt_12(self):
        refs = [Reference(text=f"Ref {i} (2020).") for i in range(5)]
        doc_content = DocumentContent(
            word_count=1000, char_count=5000, references=refs, paragraphs=["x"]
        )
        self.assertFalse(self.classifier._signal_reference_count(doc_content))

    def test_signal_s2a_false_when_no_references(self):
        doc_content = DocumentContent(word_count=1000, char_count=5000, paragraphs=["x"])
        self.assertFalse(self.classifier._signal_reference_count(doc_content))


class TestSignalReferenceRecency(unittest.TestCase):
    def setUp(self):
        self.classifier = _make_classifier()

    def test_signal_s2b_true_when_majority_recent(self):
        refs = [Reference(text=f"Autor, A. (2023). Título {i}. Revista.") for i in range(10)]
        doc = DocumentContent(word_count=1000, char_count=5000, references=refs, paragraphs=["x"])
        self.assertTrue(self.classifier._signal_reference_recency(doc))

    def test_signal_s2b_false_when_majority_old(self):
        refs = [Reference(text=f"Autor, A. (1990). Título {i}. Revista.") for i in range(10)]
        doc = DocumentContent(word_count=1000, char_count=5000, references=refs, paragraphs=["x"])
        self.assertFalse(self.classifier._signal_reference_recency(doc))

    def test_signal_s2b_false_when_no_references(self):
        doc = DocumentContent(word_count=1000, char_count=5000, paragraphs=["x"])
        self.assertFalse(self.classifier._signal_reference_recency(doc))


class TestSignalMethodologicalVocab(unittest.TestCase):
    def setUp(self):
        self.classifier = _make_classifier()

    def test_signal_s3_true_with_methodological_text(self):
        doc = DocumentContent(
            word_count=1000,
            char_count=5000,
            paragraphs=[
                "Se utilizó diseño cuasi-experimental con análisis estadístico.",
                "La muestra se seleccionó mediante muestreo teórico.",
                "Los datos primarios fueron validados mediante triangulación.",
                "Se realizó marco metodológico basado en cuantitativo.",
            ],
        )
        self.assertTrue(self.classifier._signal_methodological_vocab(doc))

    def test_signal_s3_false_with_opinion_text(self):
        doc = DocumentContent(
            word_count=1000,
            char_count=5000,
            paragraphs=["En mi opinión, esta situación es preocupante para todos."],
        )
        self.assertFalse(self.classifier._signal_methodological_vocab(doc))


# ---------------------------------------------------------------------------
# Classification rule tests (via mock LLM)
# ---------------------------------------------------------------------------


class TestClassificationRuleCientifico(unittest.TestCase):
    """Cases 2-5 should yield CIENTÍFICO."""

    def test_case2_all_signals_scientific_090(self):
        """case 2: S3+S4+S5+S2a+S2b+S6 → 0.90"""
        classifier = _make_classifier("S4: SI\nS5: SI\nS6: SI")
        doc = _scientific_doc(references=15, recent_years=True, methodological=True)
        result = classifier.classify_article(doc)
        self.assertEqual(result.article_type, ArticleType.CIENTIFICO)
        self.assertIsNotNone(result.confidence)
        self.assertGreaterEqual(result.confidence, 0.83)

    def test_case19_no_signals_opinion(self):
        """case 19: no signals → OPINIÓN"""
        classifier = _make_classifier("S4: NO\nS5: NO\nS6: NO")
        doc = DocumentContent(
            word_count=500,
            char_count=3000,
            paragraphs=["En mi opinión, el tema es relevante."],
            references=[],
        )
        result = classifier.classify_article(doc)
        self.assertEqual(result.article_type, ArticleType.OPINION)
        self.assertIsNone(result.confidence)

    def test_divulgacion_when_s3_s4_no_s5(self):
        """case 10: S3+S4, no S5 → DIVULGACION"""
        classifier = _make_classifier("S4: SI\nS5: NO\nS6: NO")
        doc = DocumentContent(
            word_count=3000,
            char_count=18000,
            paragraphs=[
                "Se utilizó metodología cuantitativa con análisis estadístico.",
                "La hipótesis fue evaluada mediante diseño cuasi-experimental.",
                "Este estudio analiza los efectos de la política educativa.",
                "Los datos primarios fueron recopilados con triangulación.",
            ],
            references=[],
        )
        result = classifier.classify_article(doc)
        self.assertIn(result.article_type, [ArticleType.DIVULGACION, ArticleType.CIENTIFICO])


class TestClassificationValidatesInputs(unittest.TestCase):
    def test_empty_document_raises(self):
        classifier = _make_classifier()
        with self.assertRaises(ValueError):
            classifier.classify_article(DocumentContent(word_count=0, char_count=0, paragraphs=[]))

    def test_none_document_raises(self):
        classifier = _make_classifier()
        with self.assertRaises((ValueError, AttributeError)):
            classifier.classify_article(None)


class TestLLMSignalParsing(unittest.TestCase):
    """_signal_s4_s5_s6 parsing logic."""

    def test_all_si_returns_true_triple(self):
        classifier = _make_classifier("S4: SI\nS5: SI\nS6: SI")
        text_sample = "sample text"
        s4, s5, s6 = classifier._signal_s4_s5_s6(text_sample, "Test title")
        self.assertTrue(s4)
        self.assertTrue(s5)
        self.assertTrue(s6)

    def test_all_no_returns_false_triple(self):
        classifier = _make_classifier("S4: NO\nS5: NO\nS6: NO")
        s4, s5, s6 = classifier._signal_s4_s5_s6("sample", "Test")
        self.assertFalse(s4)
        self.assertFalse(s5)
        self.assertFalse(s6)

    def test_mixed_signals_parsed_correctly(self):
        classifier = _make_classifier("S4: SI\nS5: NO\nS6: SI")
        s4, s5, s6 = classifier._signal_s4_s5_s6("sample", "Test")
        self.assertTrue(s4)
        self.assertFalse(s5)
        self.assertTrue(s6)

    def test_llm_exception_returns_false_triple(self):
        classifier = _make_classifier()
        classifier.client.generate.side_effect = Exception("LLM unavailable")
        s4, s5, s6 = classifier._signal_s4_s5_s6("sample", "Test")
        self.assertFalse(s4)
        self.assertFalse(s5)
        self.assertFalse(s6)


class TestIMRyDOverride(unittest.TestCase):
    """Complete IMRyD structure should override LLM signals → CIENTÍFICO."""

    def test_imryd_complete_forces_scientific(self):
        classifier = _make_classifier("S4: NO\nS5: NO\nS6: NO")
        doc = DocumentContent(
            word_count=5000,
            char_count=30000,
            paragraphs=[
                "Introducción",
                "Texto introductorio.",
                "Metodología",
                "Descripción del método.",
                "Resultados",
                "Se obtuvieron resultados.",
                "Discusión",
                "Discusión de los hallazgos.",
            ],
            references=[],
        )
        result = classifier.classify_article(doc)
        self.assertEqual(result.article_type, ArticleType.CIENTIFICO)
        self.assertEqual(result.confidence, 0.95)


if __name__ == "__main__":
    unittest.main()
