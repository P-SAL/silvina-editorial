"""
Integration tests for data_access/citation_parser.py
Uses the real .docx fixture.
"""
import sys
import os
import unittest
from unittest.mock import MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

# Defensive guard: ensure win32com mocks exist before importing any source modules
if 'win32com' not in sys.modules:
    _mock_win32com_client = MagicMock()
    _mock_win32com = MagicMock()
    _mock_win32com.client = _mock_win32com_client
    sys.modules['win32com'] = _mock_win32com
    sys.modules['win32com.client'] = _mock_win32com_client
    sys.modules['pythoncom'] = MagicMock()

FIXTURE_PATH = os.path.join(os.path.dirname(__file__), 'fixtures',
                            'capacidades_razonamiento_emergente_LLMs.docx')


class TestCitationParserIntegration(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        from data_access.citation_parser import CitationParser
        cls.parser = CitationParser()

    def test_fixture_exists(self):
        self.assertTrue(os.path.exists(FIXTURE_PATH))

    def test_extract_from_docx_returns_list(self):
        citations = self.parser.extract_from_docx(FIXTURE_PATH)
        self.assertIsInstance(citations, list)

    def test_citations_have_expected_fields(self):
        """Each Citation object must have text, citation_type, location."""
        citations = self.parser.extract_from_docx(FIXTURE_PATH)
        for cit in citations[:5]:
            self.assertTrue(hasattr(cit, 'text'))
            self.assertTrue(hasattr(cit, 'citation_type'))
            self.assertTrue(hasattr(cit, 'location'))

    def test_parse_single_paragraph(self):
        """parse() fallback method works on a paragraph string."""
        text = 'Según García (2020) la inteligencia artificial es clave.'
        citations = self.parser.parse(text, paragraph_index=0)
        self.assertIsInstance(citations, list)

    def test_parse_parenthetical(self):
        text = 'El estudio demuestra (Pérez, 2022) que el rendimiento mejora.'
        citations = self.parser.parse(text, paragraph_index=1)
        texts = [c.text for c in citations]
        self.assertTrue(any('Pérez' in t for t in texts))

    def test_empty_text_returns_empty_list(self):
        citations = self.parser.parse('', paragraph_index=0)
        self.assertIsInstance(citations, list)

    def test_nonexistent_file_returns_empty_list(self):
        """extract_from_docx gracefully handles missing file."""
        citations = self.parser.extract_from_docx('/nonexistent/file.docx')
        self.assertIsInstance(citations, list)
        self.assertEqual(citations, [])


if __name__ == '__main__':
    unittest.main()
