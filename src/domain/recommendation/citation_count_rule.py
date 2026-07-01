from src.domain.dtos.recommendation_dto import RecommendationDTO
from src.domain.enums.recommendation_priority import RecommendationPriority
from src.domain.recommendation.analysis_context import AnalysisContext
from src.domain.recommendation.recommendation_rule import RecommendationRule


class CitationCountRule(RecommendationRule):
    def evaluate(self, context: AnalysisContext) -> list[RecommendationDTO]:
        if context.citations.total_citations >= context.settings.citation_count_threshold:
            return []
        return [
            RecommendationDTO(
                priority=RecommendationPriority.MEDIUM,
                message=f"Número bajo de citas ({context.citations.total_citations}). Considere ampliar el marco teórico con más referencias.",
            )
        ]
