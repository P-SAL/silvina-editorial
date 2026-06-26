from pathlib import Path
from unittest import TestCase

from data_access.content_extractor import ContentExtractor
from src.infrastructure.wirings.extract_content_use_case_wiring import ExtractContentUseCaseWiring
from src.infrastructure.wirings.read_document_use_case_wiring import ReadDocumentUseCaseWiring

DOCS = Path(__file__).parent.parent.parent / "docs" / "sample-documents"
_DOCUMENT = "1. test_Científico.docx"


class TestExtractContentParity(TestCase):
    @classmethod
    def setUpClass(cls):
        path = str(DOCS / _DOCUMENT)
        paragraphs = ReadDocumentUseCaseWiring().create_use_case().execute(path)
        cls.legacy = ContentExtractor().extract_content(paragraphs)
        cls.result = ExtractContentUseCaseWiring().create_use_case().execute(paragraphs)

    def test_title_matches_legacy(self):
        self.assertEqual(self.result.title, self.legacy.title)

    def test_abstract_matches_legacy(self):
        self.assertEqual(self.result.abstract, self.legacy.abstract)

    def test_keywords_match_legacy(self):
        self.assertEqual(self.result.keywords, self.legacy.keywords)

    def test_sections_match_legacy(self):
        self.assertEqual(self.result.sections, self.legacy.sections)
