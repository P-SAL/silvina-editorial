from dataclasses import dataclass

from src.domain.dtos.apa_validation_result_dto import ApaValidationResultDTO
from src.domain.dtos.citation_analysis_result_dto import CitationAnalysisResultDTO
from src.domain.dtos.classification_result_dto import ClassificationResultDTO
from src.domain.dtos.grammar_check_result_dto import GrammarCheckResultDTO
from src.domain.dtos.quality_result_dto import QualityResultDTO
from src.domain.dtos.recommendation_settings_dto import RecommendationSettingsDTO
from src.domain.dtos.structure_validation_result_dto import StructureValidationResultDTO


@dataclass(frozen=True)
class AnalysisContext:
    """Groups all analysis results and settings needed to evaluate recommendations."""

    classification: ClassificationResultDTO
    quality: QualityResultDTO
    structure: StructureValidationResultDTO
    citations: CitationAnalysisResultDTO
    apa_validation: ApaValidationResultDTO
    grammar: GrammarCheckResultDTO
    settings: RecommendationSettingsDTO

    @property
    def citation_match_rate(self) -> float:
        if self.citations.total_citations == 0:
            return 100.0
        return self.citations.matched_count / self.citations.total_citations * 100.0
