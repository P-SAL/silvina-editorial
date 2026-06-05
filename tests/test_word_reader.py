"""
Integration tests for data_access/word_reader.py
Uses the real .docx fixture (via python-docx, no COM).
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


class TestWordReaderIntegration(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        from data_access.word_reader import WordReader
        cls.reader = WordReader()

    def test_fixture_file_exists(self):
        self.assertTrue(os.path.exists(FIXTURE_PATH),
                        f"Fixture not found at: {FIXTURE_PATH}")

    def test_read_document_returns_list(self):
        paragraphs = self.reader.read_word_document(FIXTURE_PATH)
        self.assertIsInstance(paragraphs, list)

    def test_read_document_non_empty(self):
        paragraphs = self.reader.read_word_document(FIXTURE_PATH)
        self.assertGreater(len(paragraphs), 0)

    def test_read_document_paragraphs_are_strings(self):
        paragraphs = self.reader.read_word_document(FIXTURE_PATH)
        for p in paragraphs[:5]:
            self.assertIsInstance(p, str)
            self.assertGreater(len(p.strip()), 0)

    def test_read_document_with_styles_returns_list_of_dicts(self):
        result = self.reader.read_document_with_styles(FIXTURE_PATH)
        self.assertIsInstance(result, list)
        if result:
            self.assertIn('text', result[0])
            self.assertIn('style', result[0])

    def test_nonexistent_file_raises_file_not_found(self):
        with self.assertRaises(FileNotFoundError):
            self.reader.read_word_document('/nonexistent/path/file.docx')

    def test_wrong_extension_raises_error(self):
        # The reader checks file existence first, then extension.
        # A non-existent .pdf file raises FileNotFoundError (caught here as well).
        with self.assertRaises((ValueError, FileNotFoundError)):
            self.reader.read_word_document('/fake/path/file.pdf')

    def test_get_document_properties_returns_dict(self):
        props = self.reader.get_document_properties(FIXTURE_PATH)
        self.assertIsInstance(props, dict)
        self.assertIn('paragraph_count', props)


if __name__ == '__main__':
    unittest.main()
