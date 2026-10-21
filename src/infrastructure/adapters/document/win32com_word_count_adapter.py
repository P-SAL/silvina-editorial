import logging
import os
from collections.abc import Generator
from contextlib import contextmanager, suppress
from typing import Any

try:
    import pythoncom
    import win32com.client

    WIN32COM_AVAILABLE = True
except ImportError:
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    WIN32COM_AVAILABLE = False

from src.domain.document.character_count_port import CharacterCountPort
from src.domain.dtos.character_count_dto import CharacterCountDTO
from src.domain.exceptions.count_errors import CharacterCountUnavailable

logger = logging.getLogger(__name__)


class Win32ComWordCountAdapter(CharacterCountPort):
    """Gets accurate character, word, and paragraph counts via COM automation."""

    def count(self, docx_path: str) -> CharacterCountDTO | None:
        """Return accurate counts for the document at the given path, or None if unavailable.

        Returns None when win32com is not installed, the file does not exist,
        or Word raises a COM error during measurement.
        """
        if not WIN32COM_AVAILABLE:
            return None
        if not os.path.exists(docx_path):
            return None
        try:
            return self._measure(path=docx_path)
        except Exception as exc:
            logger.warning("Could not get accurate Word counts: %s", exc)
            raise CharacterCountUnavailable() from exc

    def _measure(self, path: str) -> CharacterCountDTO:
        with self._word_session(path=path) as doc:
            return CharacterCountDTO(
                word_count=self._word_count(doc=doc),
                char_count=self._char_count(doc=doc),
                paragraph_count=doc.ComputeStatistics(4),
            )

    @contextmanager
    def _word_session(self, path: str) -> Generator[Any, None, None]:
        """Open a Word document via COM and yield it; always closes the document and quits Word on exit."""
        pythoncom.CoInitialize()
        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.DisplayAlerts = 0
        with suppress(Exception):
            word_app.Visible = False
        doc = None
        try:
            doc = word_app.Documents.Open(
                os.path.abspath(path),
                ReadOnly=True,
                AddToRecentFiles=False,
                ConfirmConversions=False,
            )
            yield doc
        finally:
            if doc is not None:
                doc.Close(False)
            word_app.Quit()
            pythoncom.CoUninitialize()

    def _word_count(self, doc) -> int:
        total = doc.ComputeStatistics(0)
        total += sum(fn.Range.ComputeStatistics(0) for fn in doc.Footnotes)
        total += sum(en.Range.ComputeStatistics(0) for en in doc.Endnotes)
        return total

    def _char_count(self, doc) -> int:
        total = doc.Characters.Count
        total += sum(len(fn.Range.Text) for fn in doc.Footnotes)
        total += sum(len(en.Range.Text) for en in doc.Endnotes)
        return total
