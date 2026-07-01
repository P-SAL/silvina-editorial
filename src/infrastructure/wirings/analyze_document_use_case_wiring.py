from src.application.analyze_document_use_case import AnalyzeDocumentUseCase
from src.application.analyze_quality_use_case import AnalyzeQualityUseCase
from src.application.check_grammar_use_case import CheckGrammarUseCase
from src.application.classify_article_use_case import ClassifyArticleUseCase
from src.application.extract_citations_use_case import ExtractCitationsUseCase
from src.application.extract_content_use_case import ExtractContentUseCase
from src.application.match_citations_use_case import MatchCitationsUseCase
from src.application.read_document_use_case import ReadDocumentUseCase
from src.application.validate_apa_use_case import ValidateApaUseCase
from src.application.validate_structure_use_case import ValidateStructureUseCase
from src.application.verify_eumic_use_case import VerifyEumicUseCase
from src.domain.recommendation.recommendation_builder import RecommendationBuilder
from src.infrastructure.config.recommendation_config import RecommendationConfig
from src.infrastructure.wirings.analyze_quality_use_case_wiring import AnalyzeQualityUseCaseWiring
from src.infrastructure.wirings.check_grammar_use_case_wiring import CheckGrammarUseCaseWiring
from src.infrastructure.wirings.classify_article_use_case_wiring import ClassifyArticleUseCaseWiring
from src.infrastructure.wirings.extract_citations_use_case_wiring import (
    ExtractCitationsUseCaseWiring,
)
from src.infrastructure.wirings.extract_content_use_case_wiring import ExtractContentUseCaseWiring
from src.infrastructure.wirings.match_citations_use_case_wiring import MatchCitationsUseCaseWiring
from src.infrastructure.wirings.read_document_use_case_wiring import ReadDocumentUseCaseWiring
from src.infrastructure.wirings.validate_apa_wiring import ValidateApaWiring
from src.infrastructure.wirings.validate_structure_wiring import ValidateStructureWiring
from src.infrastructure.wirings.verify_eumic_use_case_wiring import VerifyEumicUseCaseWiring


class AnalyzeDocumentUseCaseWiring:
    """Factory for building a fully wired AnalyzeDocumentUseCase."""

    def create_use_case(self) -> AnalyzeDocumentUseCase:
        return AnalyzeDocumentUseCase(
            read_document_use_case=self._get_read_document_use_case(),
            extract_content_use_case=self._get_extract_content_use_case(),
            extract_citations_use_case=self._get_extract_citations_use_case(),
            validate_apa_use_case=self._get_validate_apa_use_case(),
            check_grammar_use_case=self._get_check_grammar_use_case(),
            classify_article_use_case=self._get_classify_article_use_case(),
            analyze_quality_use_case=self._get_analyze_quality_use_case(),
            validate_structure_use_case=self._get_validate_structure_use_case(),
            match_citations_use_case=self._get_match_citations_use_case(),
            verify_eumic_use_case=self._get_verify_eumic_use_case(),
            recommendation_builder=self._get_recommendation_builder(),
        )

    def _get_read_document_use_case(self) -> ReadDocumentUseCase:
        return ReadDocumentUseCaseWiring().create_use_case()

    def _get_extract_content_use_case(self) -> ExtractContentUseCase:
        return ExtractContentUseCaseWiring().create_use_case()

    def _get_extract_citations_use_case(self) -> ExtractCitationsUseCase:
        return ExtractCitationsUseCaseWiring().create_use_case()

    def _get_validate_apa_use_case(self) -> ValidateApaUseCase:
        return ValidateApaWiring().create_use_case()

    def _get_check_grammar_use_case(self) -> CheckGrammarUseCase:
        return CheckGrammarUseCaseWiring().create_use_case()

    def _get_classify_article_use_case(self) -> ClassifyArticleUseCase:
        return ClassifyArticleUseCaseWiring().create_use_case()

    def _get_analyze_quality_use_case(self) -> AnalyzeQualityUseCase:
        return AnalyzeQualityUseCaseWiring().create_use_case()

    def _get_validate_structure_use_case(self) -> ValidateStructureUseCase:
        return ValidateStructureWiring().create_use_case()

    def _get_match_citations_use_case(self) -> MatchCitationsUseCase:
        return MatchCitationsUseCaseWiring().create_use_case()

    def _get_verify_eumic_use_case(self) -> VerifyEumicUseCase:
        return VerifyEumicUseCaseWiring().create_use_case()

    def _get_recommendation_builder(self) -> RecommendationBuilder:
        return RecommendationBuilder(settings=RecommendationConfig.build_settings())
