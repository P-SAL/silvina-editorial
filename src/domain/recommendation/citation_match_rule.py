from src.domain.dtos.recommendation_dto import RecommendationDTO
from src.domain.enums.recommendation_priority import RecommendationPriority
from src.domain.recommendation.analysis_context import AnalysisContext
from src.domain.recommendation.recommendation_rule import RecommendationRule


class CitationMatchRule(RecommendationRule):
    def evaluate(self, context: AnalysisContext) -> list[RecommendationDTO]:
        match_rate = context.citation_match_rate
        unmatched_string = "; ".join(context.citations.unmatched_citations[:10])

        if match_rate < context.settings.citation_match_threshold:
            return [
                RecommendationDTO(
                    priority=RecommendationPriority.HIGH,
                    message=f"Tasa de coincidencia de citas baja ({match_rate:.1f}%). {context.citations.unmatched_count} citas no tienen referencia correspondiente. Citas sin referencia: {unmatched_string}",
                )
            ]
        if context.citations.unmatched_count > 0:
            return [
                RecommendationDTO(
                    priority=RecommendationPriority.MEDIUM,
                    message=f"{context.citations.unmatched_count} citas no tienen referencia correspondiente. Citas sin referencia: {unmatched_string}",
                )
            ]
        return []
