from src.application.match_citations_use_case import MatchCitationsUseCase
from src.domain.citation.citation_matcher import CitationMatcher


class MatchCitationsUseCaseWiring:
    """Factory for building a ready-to-use MatchCitationsUseCase."""

    def create_use_case(self) -> MatchCitationsUseCase:
        return MatchCitationsUseCase(matcher=self._get_citation_matcher())

    def _get_citation_matcher(self) -> CitationMatcher:
        return CitationMatcher()
