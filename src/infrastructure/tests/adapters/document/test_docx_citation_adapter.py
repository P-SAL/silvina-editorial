from pathlib import Path
from unittest import TestCase

from src.domain.dtos.citation_dto import CitationDTO
from src.domain.enums.citation_type import CitationType
from src.infrastructure.adapters.document.docx_citation_adapter import DocxCitationAdapter
from src.infrastructure.adapters.document.docx_text_adapter import DocxTextAdapter

SAMPLE_DOCUMENT = (
    Path(__file__).parent.parent.parent.parent.parent.parent
    / "docs"
    / "sample-documents"
    / "1. test_Científico.docx"
)


class TestDocxCitationAdapter(TestCase):
    def setUp(self):
        self.adapter = DocxCitationAdapter(document_text_port=DocxTextAdapter())
        self.result = self.adapter.extract_citations(docx_path=str(SAMPLE_DOCUMENT))

    def test_s7a_returns_non_empty_list(self):
        self.assertIsInstance(self.result, list)
        self.assertGreater(len(self.result), 0)

    def test_s7b_all_items_are_citation_dto(self):
        for item in self.result:
            self.assertIsInstance(item, CitationDTO)

    def test_s7c_citation_type_is_author_year(self):
        for item in self.result:
            self.assertEqual(item.citation_type, CitationType.AUTHOR_YEAR)
