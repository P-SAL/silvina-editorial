"""
Integration tests for data_access/reference_parser.py
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


class TestReferenceParserIntegration(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        from data_access.reference_parser import ReferenceParser
        cls.parser = ReferenceParser()

    def test_fixture_exists(self):
        self.assertTrue(os.path.exists(FIXTURE_PATH))

    def test_parse_from_docx_returns_tuple(self):
        result = self.parser.parse_from_docx(FIXTURE_PATH)
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 2)

    def test_references_is_list(self):
        references, section_type = self.parser.parse_from_docx(FIXTURE_PATH)
        self.assertIsInstance(references, list)

    def test_section_type_is_string(self):
        references, section_type = self.parser.parse_from_docx(FIXTURE_PATH)
        self.assertIsInstance(section_type, str)

    def test_references_have_text_field(self):
        references, _ = self.parser.parse_from_docx(FIXTURE_PATH)
        for ref in references[:5]:
            self.assertTrue(hasattr(ref, 'text'))
            self.assertIsInstance(ref.text, str)

    def test_nonexistent_file_returns_empty(self):
        references, section_type = self.parser.parse_from_docx('/nonexistent/path.docx')
        self.assertIsInstance(references, list)
        self.assertEqual(references, [])

    def test_parse_section_from_text(self):
        """Compatibility method parse_section works with plain text."""
        bib_text = (
            'García, J. (2020). Título del artículo. Revista Académica, 1(1), 1-10.\n'
            'López, M. (2021). Otro artículo. Journal, 2(2), 5-15.'
        )
        references, section_type = self.parser.parse_section(bib_text)
        self.assertIsInstance(references, list)

    def test_parse_section_none_returns_empty(self):
        references, section_type = self.parser.parse_section(None)
        self.assertEqual(references, [])


if __name__ == '__main__':
    unittest.main()
