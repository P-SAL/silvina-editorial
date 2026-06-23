from src.domain.classification.classification_specification import ClassificationSpecification
from src.domain.dtos.classification_signals_dto import ClassificationSignalsDTO


class HasTheoreticalJustificationSpecification(ClassificationSpecification):
    """Satisfied when theoretical justification is detected (legacy signal S6)."""

    def is_satisfied_by(self, signals: ClassificationSignalsDTO) -> bool:
        return signals.has_theoretical_justification
