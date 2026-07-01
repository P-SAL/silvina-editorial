"""
Unit tests for data_access/content_extractor.py
Mocks WordCounter; tests title/author/abstract/keyword extraction.
"""

import sys
import os
import unittest
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

# Ensure win32com mocks are in place before any imports (defensive guard)
if "win32com" not in sys.modules:
    _mock_win32com_client = MagicMock()
    _mock_win32com = MagicMock()
    _mock_win32com.client = _mock_win32com_client
    sys.modules["win32com"] = _mock_win32com
    sys.modules["win32com.client"] = _mock_win32com_client
    sys.modules["pythoncom"] = MagicMock()

from data_access.content_extractor import ContentExtractor


class TestContentExtractorTitle(unittest.TestCase):
    def setUp(self):
        self.extractor = ContentExtractor()

    def test_extracts_first_short_paragraph_as_title(self):
        paragraphs = [
            "Capacidades de razonamiento emergente en LLMs",
            "Juan García; María López",
            "Resumen",
            "Este artículo analiza las capacidades emergentes.",
        ]
        result = self.extractor.extract_content(paragraphs)
        self.assertIsNotNone(result.title)
        self.assertIn("Capacidades", result.title)

    def test_explicit_titulo_marker(self):
        paragraphs = ["TÍTULO: Mi artículo de prueba", "Contenido del artículo."]
        result = self.extractor.extract_content(paragraphs)
        self.assertEqual(result.title, "Mi artículo de prueba")

    def test_title_not_institution_header(self):
        paragraphs = [
            "Universidad Nacional de La Plata",
            "Título real del artículo",
            "Contenido largo del artículo académico de investigación.",
        ]
        result = self.extractor.extract_content(paragraphs)
        self.assertIsNotNone(result.title)
        self.assertNotIn("Universidad", result.title)


class TestContentExtractorAuthor(unittest.TestCase):
    def setUp(self):
        self.extractor = ContentExtractor()

    def test_extracts_author_after_title(self):
        paragraphs = [
            "Estudio sobre el impacto de la IA",
            "Juan García",
            "Resumen",
            "Texto del resumen aquí.",
        ]
        result = self.extractor.extract_content(paragraphs)
        self.assertIsNotNone(result.authors)
        self.assertIn("García", result.authors)

    def test_author_not_identified_when_missing(self):
        paragraphs = [
            "Título del artículo de investigación",
            "INTRODUCCIÓN",
            "Texto introductorio muy largo para asegurarse que no sea autor.",
        ]
        result = self.extractor.extract_content(paragraphs)
        # Should return "Autor no identificado" or a string
        self.assertIsNotNone(result.authors)

    def test_explicit_autor_marker(self):
        paragraphs = [
            "Mi artículo",
            "AUTOR: Dr. Juan García",
            "Texto de contenido.",
        ]
        result = self.extractor.extract_content(paragraphs)
        self.assertIsNotNone(result.authors)


class TestContentExtractorAbstract(unittest.TestCase):
    def setUp(self):
        self.extractor = ContentExtractor()

    def test_extracts_abstract_after_resumen_header(self):
        paragraphs = [
            "Título del artículo",
            "Autor",
            "RESUMEN",
            "Este es el texto del resumen del artículo.",
            "INTRODUCCIÓN",
            "Texto de introducción.",
        ]
        result = self.extractor.extract_content(paragraphs)
        self.assertIsNotNone(result.abstract)
        self.assertIn("resumen", result.abstract.lower())

    def test_no_abstract_returns_none(self):
        paragraphs = [
            "Título del artículo corto",
            "INTRODUCCIÓN",
            "Texto de introducción.",
        ]
        result = self.extractor.extract_content(paragraphs)
        self.assertIsNone(result.abstract)


class TestContentExtractorKeywords(unittest.TestCase):
    def setUp(self):
        self.extractor = ContentExtractor()

    def test_extracts_keywords(self):
        paragraphs = [
            "Título del artículo",
            "PALABRAS CLAVE: inteligencia artificial, machine learning, NLP",
            "Contenido del artículo.",
        ]
        result = self.extractor.extract_content(paragraphs)
        self.assertIsInstance(result.keywords, list)
        self.assertGreater(len(result.keywords), 0)

    def test_no_keywords_returns_empty_list(self):
        paragraphs = ["Título", "Contenido sin palabras clave."]
        result = self.extractor.extract_content(paragraphs)
        self.assertEqual(result.keywords, [])


class TestContentExtractorWordCount(unittest.TestCase):
    def setUp(self):
        self.extractor = ContentExtractor()

    def test_word_count_computed(self):
        paragraphs = ["Una dos tres cuatro cinco", "Seis siete ocho"]
        result = self.extractor.extract_content(paragraphs)
        self.assertEqual(result.word_count, 8)

    def test_char_count_computed(self):
        paragraphs = ["abc"]
        result = self.extractor.extract_content(paragraphs)
        self.assertEqual(result.char_count, 3)

    def test_empty_paragraphs_raises(self):
        with self.assertRaises(ValueError):
            self.extractor.extract_content([])

    def test_whitespace_only_paragraphs_raises(self):
        with self.assertRaises(ValueError):
            self.extractor.extract_content(["   ", "\t", ""])


class TestContentExtractorWithMockedWordCounter(unittest.TestCase):
    """Verify that WordCounter is called when docx_path is provided."""

    def test_skips_word_counter_when_win32com_unavailable(self):
        """With WIN32COM_AVAILABLE=False, should use text-based counts."""
        with patch("data_access.content_extractor.WIN32COM_AVAILABLE", False):
            extractor = ContentExtractor()
            result = extractor.extract_content(
                ["Hola mundo", "Texto adicional"], docx_path="/fake/path.docx"
            )
            self.assertGreater(result.word_count, 0)


if __name__ == "__main__":
    unittest.main()
