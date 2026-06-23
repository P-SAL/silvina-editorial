from src.domain.classification.classification_specification import ClassificationSpecification
from src.domain.dtos.classification_signals_dto import ClassificationSignalsDTO


class HasRecentReferencesSpecification(ClassificationSpecification):
    """Satisfied when most references are recent (legacy signal S2b)."""

    def is_satisfied_by(self, signals: ClassificationSignalsDTO) -> bool:
        return signals.has_recent_references
