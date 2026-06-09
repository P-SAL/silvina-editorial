"""
Unit tests for business_logic/structure_analyzer.py
Tests IMRyD section boundary detection using mock DocumentContent.
"""
import sys
import os
import unittest
from unittest.mock import MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

# Defensive guard for COM mocks
if 'win32com' not in sys.modules:
    _wc = MagicMock(); _wcc = MagicMock(); _wc.client = _wcc
    sys.modules.update({'win32com': _wc, 'win32com.client': _wcc, 'pythoncom': MagicMock()})

from business_logic.structure_analyzer import StructureAnalyzer, analyze_structure
from domain.models import DocumentContent


def _doc(paragraphs):
    return DocumentContent(word_count=1000, char_count=5000, paragraphs=paragraphs)


class TestStructureAnalyzerIMRyD(unittest.TestCase):

    def setUp(self):
        self.analyzer = StructureAnalyzer()

    def test_full_imryd_complete(self):
        doc = _doc([
            'Introducción',
            'Texto introductorio largo.',
            'Metodología',
            'Se aplicó encuesta.',
            'Resultados',
            'Los resultados muestran mejora.',
            'Discusión',
            'Estos resultados implican X.',
            'Conclusiones',
            'En conclusión Y.',
        ])
        result = self.analyzer.analyze(doc)
        self.assertTrue(result['imryd_complete'])
        self.assertTrue(result['has_introduction'])
        self.assertTrue(result['has_methods'])
        self.assertTrue(result['has_results'])
        self.assertTrue(result['has_discussion'])

    def test_missing_methods_not_complete(self):
        doc = _doc([
            'Introducción',
            'Resultados',
            'Discusión',
        ])
        result = self.analyzer.analyze(doc)
        self.assertFalse(result['imryd_complete'])
        self.assertFalse(result['has_methods'])

    def test_english_keywords_detected(self):
        doc = _doc([
            'Introduction',
            'Methodology',
            'Results',
            'Discussion',
        ])
        result = self.analyzer.analyze(doc)
        self.assertTrue(result['imryd_complete'])

    def test_body_prose_not_false_positive(self):
        """Body paragraphs > 5 words should NOT trigger section detection."""
        doc = _doc([
            'En esta sección de introducción se explica el contexto del estudio.',
            'La metodología empleada fue cuantitativa y se basa en encuestas.',
            'Los resultados obtenidos demuestran que la hipótesis es correcta.',
            'La discusión aborda las limitaciones y el alcance del trabajo.',
        ])
        result = self.analyzer.analyze(doc)
        # Long paragraphs should not match section headers
        self.assertFalse(result['imryd_complete'])

    def test_no_sections_all_false(self):
        doc = _doc(['Texto sin ninguna sección formal detectada en este documento.'])
        result = self.analyzer.analyze(doc)
        self.assertFalse(result['has_introduction'])
        self.assertFalse(result['has_methods'])
        self.assertFalse(result['has_results'])
        self.assertFalse(result['has_discussion'])
        self.assertFalse(result['imryd_complete'])

    def test_analyze_returns_all_signal_keys(self):
        doc = _doc([])
        result = self.analyzer.analyze(doc)
        expected_keys = {
            'has_introduction', 'has_methods', 'has_results',
            'has_discussion', 'has_conclusion', 'imryd_complete'
        }
        self.assertEqual(set(result.keys()), expected_keys)

    def test_convenience_function(self):
        doc = _doc(['Introducción', 'Metodología', 'Resultados', 'Discusión'])
        result = analyze_structure(doc)
        self.assertIsInstance(result, dict)
        self.assertTrue(result['imryd_complete'])

    def test_materiales_y_metodos_variant(self):
        doc = _doc([
            'Introducción',
            'Materiales y métodos',
            'Resultados',
            'Discusión',
        ])
        result = self.analyzer.analyze(doc)
        # "materiales" is in methods keywords but the paragraph has 3 words — within limit
        self.assertTrue(result['has_methods'])

    def test_spanish_conclusion_detected(self):
        doc = _doc(['Conclusiones'])
        result = self.analyzer.analyze(doc)
        self.assertTrue(result['has_conclusion'])


if __name__ == '__main__':
    unittest.main()
