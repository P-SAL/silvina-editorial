"""
Smoke test: parity between legacy WordReader and new ReadDocumentUseCase.

Legacy path:
    WordReader.read_word_document(path) -> list[str]

New path:
    ReadDocumentUseCaseWiring().create_use_case().execute(path) -> list[str]

Run with: python -m pytest tests/smoke/ -v
"""

from pathlib import Path
from unittest import TestCase

from data_access.word_reader import WordReader
from src.infrastructure.wirings.analyze_document_use_case_wiring import AnalyzeDocumentUseCaseWiring

DOCS = Path(__file__).parent.parent.parent / "docs" / "sample-documents"

_DOCUMENTS = [
    "1. test_Científico.docx",
    "2. test_divulgacion_v2.docx",
    "3. test_opinion_v2.docx",
]


class TestReadDocumentParity(TestCase):
    @classmethod
    def setUpClass(cls):
        cls.legacy_reader = WordReader()
        cls.document_text_port = AnalyzeDocumentUseCaseWiring()._get_document_text_port()

    def test_cientifico_matches_legacy(self):
        path = str(DOCS / _DOCUMENTS[0])
        self.assertEqual(
            self.document_text_port.read_paragraphs(path=path),
            self.legacy_reader.read_word_document(path),
        )

    def test_divulgacion_matches_legacy(self):
        path = str(DOCS / _DOCUMENTS[1])
        self.assertEqual(
            self.document_text_port.read_paragraphs(path=path),
            self.legacy_reader.read_word_document(path),
        )

    def test_opinion_matches_legacy(self):
        path = str(DOCS / _DOCUMENTS[2])
        self.assertEqual(
            self.document_text_port.read_paragraphs(path=path),
            self.legacy_reader.read_word_document(path),
        )
