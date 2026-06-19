from unittest import TestCase

from src.application.match_citations_use_case import MatchCitationsUseCase
from src.domain.dtos.citation_analysis_result_dto import CitationAnalysisResultDTO
from src.domain.enums.section_name import SectionName
from src.infrastructure.wirings.match_citations_use_case_wiring import MatchCitationsUseCaseWiring


class TestMatchCitationsUseCaseWiring(TestCase):
    def setUp(self):
        self.wiring = MatchCitationsUseCaseWiring()

    def test_s32_create_use_case_returns_correct_type(self):
        use_case = self.wiring.create_use_case()
        self.assertIsInstance(use_case, MatchCitationsUseCase)

    def test_s33_use_case_execute_returns_citation_analysis_result(self):
        use_case = self.wiring.create_use_case()
        result = use_case.execute([], [], SectionName.REFERENCES)
        self.assertIsInstance(result, CitationAnalysisResultDTO)
