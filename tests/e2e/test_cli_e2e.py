"""
E2E test for main.py SilvinaEditorialAssistant orchestrator.
Runs in-process with mocked external dependencies (Ollama, LanguageTool, COM).
"""
import sys
import os
import unittest
from unittest.mock import MagicMock, patch

# Path adjustment from tests/e2e/ → project root
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

# Inject language_tool_python mock before any import that might trigger it
if 'language_tool_python' not in sys.modules:
    _mock_ltp = MagicMock()
    _mock_lt_instance = MagicMock()
    _mock_lt_instance.check.return_value = []
    _mock_ltp.LanguageTool.return_value = _mock_lt_instance
    sys.modules['language_tool_python'] = _mock_ltp

FIXTURE_PATH = os.path.join(
    os.path.dirname(__file__), '..', 'fixtures',
    'capacidades_razonamiento_emergente_LLMs.docx'
)

FIXTURE_PATH = os.path.abspath(FIXTURE_PATH)


def _make_ollama_client_mock(response='S4: SI\nS5: SI\nS6: SI'):
    """Return a mock ollama.Client whose generate() returns a fixed response."""
    mock_client = MagicMock()
    mock_client.generate.return_value = {'response': response}
    return mock_client


class TestCLIE2E(unittest.TestCase):
    """
    In-process E2E test: constructs SilvinaEditorialAssistant, calls
    analyze_document() with a real .docx fixture, verifies the result dict.
    """

    @classmethod
    def setUpClass(cls):
        cls.fixture_exists = os.path.exists(FIXTURE_PATH)

    def test_fixture_available(self):
        self.assertTrue(
            self.fixture_exists,
            f"Fixture not found at: {FIXTURE_PATH}"
        )

    @unittest.skipUnless(
        os.path.exists(FIXTURE_PATH),
        "Fixture .docx not available"
    )
    def test_analyze_document_returns_dict(self):
        mock_client = _make_ollama_client_mock()
        with patch('ollama.Client', return_value=mock_client), \
             patch('data_access.word_counter.WIN32COM_AVAILABLE', False):
            from main import SilvinaEditorialAssistant
            silvina = SilvinaEditorialAssistant()
            result = silvina.analyze_document(FIXTURE_PATH)
        self.assertIsInstance(result, dict)

    @unittest.skipUnless(
        os.path.exists(FIXTURE_PATH),
        "Fixture .docx not available"
    )
    def test_result_has_required_keys(self):
        mock_client = _make_ollama_client_mock()
        with patch('ollama.Client', return_value=mock_client), \
             patch('data_access.word_counter.WIN32COM_AVAILABLE', False):
            from main import SilvinaEditorialAssistant
            silvina = SilvinaEditorialAssistant()
            result = silvina.analyze_document(FIXTURE_PATH)
        for key in ['filename', 'document_info', 'classification',
                    'quality_analysis', 'structure_validation']:
            self.assertIn(key, result)

    @unittest.skipUnless(
        os.path.exists(FIXTURE_PATH),
        "Fixture .docx not available"
    )
    def test_classification_key_has_category(self):
        mock_client = _make_ollama_client_mock()
        with patch('ollama.Client', return_value=mock_client), \
             patch('data_access.word_counter.WIN32COM_AVAILABLE', False):
            from main import SilvinaEditorialAssistant
            silvina = SilvinaEditorialAssistant()
            result = silvina.analyze_document(FIXTURE_PATH)
        self.assertIn('category', result['classification'])

    @unittest.skipUnless(
        os.path.exists(FIXTURE_PATH),
        "Fixture .docx not available"
    )
    def test_document_info_has_word_count(self):
        mock_client = _make_ollama_client_mock()
        with patch('ollama.Client', return_value=mock_client), \
             patch('data_access.word_counter.WIN32COM_AVAILABLE', False):
            from main import SilvinaEditorialAssistant
            silvina = SilvinaEditorialAssistant()
            result = silvina.analyze_document(FIXTURE_PATH)
        self.assertGreater(result['document_info']['word_count'], 0)


if __name__ == '__main__':
    unittest.main()
