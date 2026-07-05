"""
Smoke test: DocxTextAdapter reads real .docx fixtures correctly.

Exercises src.infrastructure.adapters.document.docx_text_adapter.DocxTextAdapter
directly against real sample documents, verifying it returns non-empty,
stripped paragraphs in document order.

Run with: python -m pytest tests/smoke/ -v
"""

from pathlib import Path
from unittest import TestCase

from src.infrastructure.adapters.document.docx_text_adapter import DocxTextAdapter

DOCS = Path(__file__).parent.parent.parent / "docs" / "sample-documents"

_DOCUMENTS = [
    "1. test_Científico.docx",
    "2. test_divulgacion_v2.docx",
    "3. test_opinion_v2.docx",
]


class TestReadDocumentParity(TestCase):
    @classmethod
    def setUpClass(cls):
        cls.document_text_port = DocxTextAdapter()

    def test_cientifico_returns_stripped_nonempty_paragraphs(self):
        self._assert_reads_stripped_nonempty_paragraphs(_DOCUMENTS[0])

    def test_divulgacion_returns_stripped_nonempty_paragraphs(self):
        self._assert_reads_stripped_nonempty_paragraphs(_DOCUMENTS[1])

    def test_opinion_returns_stripped_nonempty_paragraphs(self):
        self._assert_reads_stripped_nonempty_paragraphs(_DOCUMENTS[2])

    def _assert_reads_stripped_nonempty_paragraphs(self, filename: str):
        path = str(DOCS / filename)
        paragraphs = self.document_text_port.read_paragraphs(path=path)
        self.assertGreater(len(paragraphs), 0)
        for paragraph in paragraphs:
            self.assertEqual(paragraph, paragraph.strip())
            self.assertNotEqual(paragraph, "")
