from unittest import TestCase
from unittest.mock import patch

from src.domain.exceptions.count_errors import CharacterCountUnavailable
from src.infrastructure.adapters.document.win32com_word_count_adapter import (
    Win32ComWordCountAdapter,
)


class TestWin32ComWordCountAdapter(TestCase):
    def setUp(self):
        self.adapter = Win32ComWordCountAdapter()

    def test_returns_none_when_win32com_unavailable(self):
        with patch(
            "src.infrastructure.adapters.document.win32com_word_count_adapter.WIN32COM_AVAILABLE",
            False,
        ):
            result = self.adapter.count(docx_path="any/path.docx")
        self.assertIsNone(result)

    def test_raises_character_count_unavailable_on_com_exception(self):
        with (
            patch(
                "src.infrastructure.adapters.document.win32com_word_count_adapter.WIN32COM_AVAILABLE",
                True,
            ),
            patch(
                "src.infrastructure.adapters.document.win32com_word_count_adapter.win32com"
            ) as mock_win32com,
            patch("os.path.exists", return_value=True),
        ):
            mock_win32com.client.DispatchEx.side_effect = Exception("COM failure")
            with self.assertRaises(CharacterCountUnavailable):
                self.adapter.count(docx_path="any/path.docx")
