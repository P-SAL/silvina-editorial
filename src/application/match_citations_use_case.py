from src.domain.citation.citation_matcher import CitationMatcher
from src.domain.dtos.citation_analysis_result_dto import CitationAnalysisResultDTO
from src.domain.dtos.citation_dto import CitationDTO
from src.domain.dtos.reference_dto import ReferenceDTO
from src.domain.enums.section_name import SectionName


class MatchCitationsUseCase:
    def __init__(self, matcher: CitationMatcher) -> None:
        self._matcher = matcher

    def execute(
        self,
        citations: list[CitationDTO],
        references: list[ReferenceDTO],
        section_type: SectionName = SectionName.REFERENCES,
    ) -> CitationAnalysisResultDTO:
        return self._matcher.match_citations_to_references(
            citations=citations, references=references, section_type=section_type
        )
