from src.domain.classification.classification_specification import ClassificationSpecification
from src.domain.dtos.classification_signals_dto import ClassificationSignalsDTO


class HasResearchIntentSpecification(ClassificationSpecification):
    """Satisfied when explicit research intent is detected (legacy signal S4)."""

    def is_satisfied_by(self, signals: ClassificationSignalsDTO) -> bool:
        return signals.has_research_intent
