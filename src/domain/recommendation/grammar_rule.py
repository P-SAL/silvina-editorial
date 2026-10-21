from src.domain.dtos.recommendation_dto import RecommendationDTO
from src.domain.enums.recommendation_priority import RecommendationPriority
from src.domain.recommendation.analysis_context import AnalysisContext
from src.domain.recommendation.recommendation_rule import RecommendationRule


class GrammarRule(RecommendationRule):
    def evaluate(self, context: AnalysisContext) -> list[RecommendationDTO]:
        if context.grammar.score >= context.settings.grammar_threshold:
            return []
        return [
            RecommendationDTO(
                priority=RecommendationPriority.HIGH,
                message=f"Gramática ({context.grammar.score:.1f}/10) requiere corrección.",
            )
        ]
