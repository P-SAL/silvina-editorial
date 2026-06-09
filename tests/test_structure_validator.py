"""
Unit tests for business_logic/structure_validator.py
"""
import unittest
import sys
import os
from unittest.mock import MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

# Defensive guard for COM mocks
if 'win32com' not in sys.modules:
    _wc = MagicMock(); _wcc = MagicMock(); _wc.client = _wcc
    sys.modules.update({'win32com': _wc, 'win32com.client': _wcc, 'pythoncom': MagicMock()})

from business_logic.structure_validator import StructureValidator
from domain.models import DocumentContent
from domain.enums import ArticleType


class TestStructureValidator(unittest.TestCase):

    def setUp(self):
        self.validator = StructureValidator()

    def test_scientific_article_detects_all_imryd_sections(self):
        """All IMRyD sections are correctly detected when present."""
        content = DocumentContent(
            word_count=1000,
            char_count=5000,
            paragraphs=[
                "Resumen: Este es el resumen.",
                "Introducción: Bienvenidos al paper.",
                "Metodología: Usamos python.",
                "Resultados: Encontramos bugs.",
                "Discusión: Esto significa que funciona.",
                "Conclusiones: Fin.",
                "Referencias: [1] Knuth, D."
            ]
        )
        result = self.validator.validate_structure(content, ArticleType.CIENTIFICO)
        self.assertTrue(result.is_valid)
        self.assertEqual(len(result.missing_sections), 0)

    def test_scientific_article_complete_is_valid(self):
        """A complete scientific article with all sections validates successfully."""
        content = DocumentContent(
            word_count=1000, char_count=5000,
            paragraphs=[
                "Resumen: El resumen.",
                "Introducción: Intro.",
                "Metodología: Método.",
                "Resultados: Los resultados.",
                "Discusión: Discusión.",
                "Conclusiones: Conclusión.",
                "Referencias: Refs.",
            ]
        )
        result = self.validator.validate_structure(content, ArticleType.CIENTIFICO)
        self.assertTrue(result.is_valid)
        self.assertEqual(result.missing_sections, [])

    def test_scientific_article_missing_resumen(self):
        content = DocumentContent(
            word_count=1000, char_count=5000,
            paragraphs=[
                "Introducción: Bienvenidos al paper.",
                "Conclusiones: Fin.",
                "Referencias: [1] Knuth, D."
            ]
        )
        result = self.validator.validate_structure(content, ArticleType.CIENTIFICO)
        self.assertFalse(result.is_valid)
        self.assertIn("Resumen", result.missing_sections)

    def test_divulgacion_article_compliant(self):
        """Divulgación only requires sections that ARE in section_map."""
        content = DocumentContent(
            word_count=1000, char_count=5000,
            paragraphs=[
                "Resumen: Resumen de divulgacion.",
                "Introducción: Introduccion de divulgacion.",
                "Desarrollo: Desarrollo de divulgacion.",
                "Conclusiones: Conclusiones de divulgacion.",
                "Referencias: Referencias."
            ]
        )
        result = self.validator.validate_structure(content, ArticleType.DIVULGACION)
        self.assertTrue(result.is_valid)
        self.assertEqual(len(result.missing_sections), 0)

    def test_divulgacion_missing_desarrollo(self):
        content = DocumentContent(
            word_count=1000, char_count=5000,
            paragraphs=[
                "Resumen: Resumen.",
                "Introducción: Intro.",
                "Conclusiones: Conclusión.",
                "Referencias: Refs.",
            ]
        )
        result = self.validator.validate_structure(content, ArticleType.DIVULGACION)
        self.assertFalse(result.is_valid)
        self.assertIn("Desarrollo", result.missing_sections)

    def test_opinion_article_complete_is_valid(self):
        """A complete opinion article with all required sections validates successfully."""
        content = DocumentContent(
            word_count=1000, char_count=5000,
            paragraphs=[
                "Introducción: Introduccion de opinion.",
                "Argumentación: Argumentos de opinion.",
                "Conclusiones: Conclusiones de opinion."
            ]
        )
        result = self.validator.validate_structure(content, ArticleType.OPINION)
        self.assertTrue(result.is_valid)
        self.assertEqual(result.missing_sections, [])

    def test_validate_structure_returns_result_object(self):
        content = DocumentContent(word_count=100, char_count=500, paragraphs=["Texto."])
        result = self.validator.validate_structure(content, ArticleType.DIVULGACION)
        self.assertTrue(hasattr(result, 'is_valid'))
        self.assertTrue(hasattr(result, 'missing_sections'))

    def test_english_abstract_detected(self):
        """'abstract' keyword maps to 'resumen' section."""
        content = DocumentContent(
            word_count=1000, char_count=5000,
            paragraphs=["Abstract: This is the abstract.", "Introducción: Intro."]
        )
        present = self.validator._extract_present_sections(content)
        present_lower = [p.lower() for p in present]
        self.assertIn('resumen', present_lower)

    def test_section_aliases_detected(self):
        """Aliases in section_map are correctly detected."""
        content = DocumentContent(
            word_count=1000, char_count=5000,
            paragraphs=[
                "metodologia: Methods without accent.",
                "methodology: English version of methods.",
                "discussion: English version of discussion.",
                "results: English results.",
            ]
        )
        present = self.validator._extract_present_sections(content)
        present_lower = [p.lower() for p in present]
        self.assertIn('metodología', present_lower)
        self.assertIn('discusión', present_lower)
        self.assertIn('resultados', present_lower)

    def test_long_body_text_not_detected_as_section(self):
        """Paragraphs >= 100 chars are not detected as section headers even if they contain keywords."""
        long_para = "La introducción de nuevas metodologías en el campo de la investigación académica requiere un análisis."
        self.assertGreaterEqual(len(long_para), 100)
        content = DocumentContent(word_count=500, char_count=3000, paragraphs=[long_para])
        present = self.validator._extract_present_sections(content)
        self.assertEqual(present, [])

    def test_short_section_header_is_detected(self):
        """Paragraphs < 100 chars with section keywords ARE detected as headers (threshold boundary)."""
        short_header = "Introducción"
        self.assertLess(len(short_header), 100)
        content = DocumentContent(word_count=100, char_count=500, paragraphs=[short_header])
        present = self.validator._extract_present_sections(content)
        self.assertIn("Introducción", present)


if __name__ == "__main__":
    unittest.main()
