from src.domain.classification.classification_specification import ClassificationSpecification
from src.domain.dtos.classification_signals_dto import ClassificationSignalsDTO


class HasSufficientReferenceCountSpecification(ClassificationSpecification):
    """Satisfied when the document has enough references (legacy signal S2a)."""

    def is_satisfied_by(self, signals: ClassificationSignalsDTO) -> bool:
        return signals.has_sufficient_reference_count
