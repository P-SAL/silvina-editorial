from unittest import TestCase

from src.application.match_citations_use_case import MatchCitationsUseCase
from src.domain.citation.citation_matcher import CitationMatcher
from src.domain.dtos.citation_dto import CitationDTO
from src.domain.dtos.reference_dto import ReferenceDTO
from src.domain.enums.citation_type import CitationType
from src.domain.enums.section_name import SectionName


class TestMatchCitationsUseCase(TestCase):
    def setUp(self):
        self.use_case = MatchCitationsUseCase(matcher=CitationMatcher())

    def test_s31_execute_matches_domain_service_result(self):
        matched = CitationDTO(
            text="(Smith, 2020)", citation_type=CitationType.AUTHOR_YEAR, location=0, author="Smith"
        )
        unmatched = CitationDTO(
            text="(Jones, 2021)", citation_type=CitationType.AUTHOR_YEAR, location=1, author="Jones"
        )
        reference = ReferenceDTO(text="Smith, J. (2020). Title.")
        citations = [matched, unmatched]
        references = [reference]

        expected = CitationMatcher().match_citations_to_references(
            citations, references, section_type=SectionName.REFERENCES
        )
        result = self.use_case.execute(citations, references, section_type=SectionName.REFERENCES)

        self.assertEqual(result.total_citations, expected.total_citations)
        self.assertEqual(result.matched_count, expected.matched_count)
        self.assertEqual(result.unmatched_count, expected.unmatched_count)
        self.assertEqual(result.unmatched_citations, expected.unmatched_citations)
