from unittest import TestCase

from src.application.extract_citations_use_case import ExtractCitationsUseCase
from src.infrastructure.adapters.document.docx_citation_adapter import DocxCitationAdapter
from src.infrastructure.adapters.document.docx_reference_adapter import DocxReferenceAdapter
from src.infrastructure.wirings.extract_citations_use_case_wiring import (
    ExtractCitationsUseCaseWiring,
)


class TestExtractCitationsUseCaseWiring(TestCase):
    def test_create_use_case_returns_use_case_with_correct_adapters(self):
        use_case = ExtractCitationsUseCaseWiring().create_use_case()

        self.assertIsInstance(use_case, ExtractCitationsUseCase)
        self.assertIsInstance(use_case._citation_port, DocxCitationAdapter)
        self.assertIsInstance(use_case._reference_port, DocxReferenceAdapter)
