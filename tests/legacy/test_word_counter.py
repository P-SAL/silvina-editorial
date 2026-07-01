"""
Unit tests for data_access/word_counter.py
Mocks win32com.client.DispatchEx to test logic paths without COM.
"""

import sys
import os
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

# Ensure win32com mocks are present (defensive — also done in tests/__init__.py)
if "win32com" not in sys.modules:
    _mock_win32com_client = MagicMock()
    _mock_win32com = MagicMock()
    _mock_win32com.client = _mock_win32com_client
    sys.modules["win32com"] = _mock_win32com
    sys.modules["win32com.client"] = _mock_win32com_client
    sys.modules["pythoncom"] = MagicMock()

from data_access.word_counter import WordCounter


class TestWordCounterWithoutWin32com(unittest.TestCase):
    """When win32com is not available, get_accurate_counts returns None."""

    def test_returns_none_when_win32com_unavailable(self):
        with patch("data_access.word_counter.WIN32COM_AVAILABLE", False):
            counter = WordCounter()
            result = counter.get_accurate_counts("/fake/path.docx")
            self.assertIsNone(result)

    def test_returns_none_for_nonexistent_file(self):
        with patch("data_access.word_counter.WIN32COM_AVAILABLE", True):
            counter = WordCounter()
            result = counter.get_accurate_counts("/nonexistent/path/file.docx")
            self.assertIsNone(result)


class TestWordCounterMockedCOM(unittest.TestCase):
    """Mock win32com.client to test the happy path."""

    def _make_word_app_mock(self):
        mock_doc = MagicMock()
        mock_doc.Characters.Count = 1000
        mock_doc.ComputeStatistics.return_value = 200
        mock_doc.Footnotes = []
        mock_doc.Endnotes = []

        mock_app = MagicMock()
        mock_app.Documents.Open.return_value = mock_doc

        return mock_app, mock_doc

    def test_happy_path_returns_counts_dict(self):
        mock_app, mock_doc = self._make_word_app_mock()

        # Directly replace DispatchEx on the injected mock, save and restore
        wc_client = sys.modules["win32com.client"]
        original_dispatch = wc_client.DispatchEx
        mock_dispatch = MagicMock(return_value=mock_app)
        wc_client.DispatchEx = mock_dispatch
        try:
            with (
                patch("data_access.word_counter.WIN32COM_AVAILABLE", True),
                patch("os.path.exists", return_value=True),
            ):
                counter = WordCounter()
                result = counter.get_accurate_counts("/fake/doc.docx")
        finally:
            wc_client.DispatchEx = original_dispatch

        self.assertIsNotNone(result)
        self.assertIn("char_count", result)
        self.assertIn("word_count", result)
        self.assertIn("paragraph_count", result)

    def test_com_exception_returns_none(self):
        """If COM raises on open, returns None after retries."""
        # Override DispatchEx on the already-injected win32com.client MagicMock
        original_dispatch = sys.modules["win32com.client"].DispatchEx
        sys.modules["win32com.client"].DispatchEx = MagicMock(side_effect=Exception("COM error"))
        try:
            with (
                patch("data_access.word_counter.WIN32COM_AVAILABLE", True),
                patch("os.path.exists", return_value=True),
                patch("time.sleep", return_value=None),
            ):
                counter = WordCounter()
                result = counter.get_accurate_counts("/fake/doc.docx")
        finally:
            sys.modules["win32com.client"].DispatchEx = original_dispatch

        self.assertIsNone(result)


class TestWordCounterHelpers(unittest.TestCase):
    """Test the private helper methods via a mocked self.doc."""

    def test_get_character_count_no_doc(self):
        counter = WordCounter()
        counter.doc = None
        self.assertEqual(counter._get_character_count(), 0)

    def test_get_word_count_no_doc(self):
        counter = WordCounter()
        counter.doc = None
        self.assertEqual(counter._get_word_count(), 0)

    def test_get_paragraph_count_no_doc(self):
        counter = WordCounter()
        counter.doc = None
        self.assertEqual(counter._get_paragraph_count(), 0)

    def test_get_character_count_with_mock_doc(self):
        counter = WordCounter()
        mock_doc = MagicMock()
        mock_doc.Characters.Count = 500
        mock_doc.Footnotes = []
        mock_doc.Endnotes = []
        counter.doc = mock_doc
        self.assertEqual(counter._get_character_count(), 500)

    def test_get_word_count_with_mock_doc(self):
        counter = WordCounter()
        mock_doc = MagicMock()
        mock_doc.ComputeStatistics.return_value = 100
        mock_doc.Footnotes = []
        mock_doc.Endnotes = []
        counter.doc = mock_doc
        self.assertEqual(counter._get_word_count(), 100)


if __name__ == "__main__":
    unittest.main()
