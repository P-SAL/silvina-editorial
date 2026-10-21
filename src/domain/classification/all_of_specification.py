from src.domain.classification.classification_specification import ClassificationSpecification
from src.domain.dtos.classification_signals_dto import ClassificationSignalsDTO


class AllOfSpecification(ClassificationSpecification):
    """Satisfied when every wrapped specification is satisfied (logical AND)."""

    def __init__(self, *specifications: ClassificationSpecification) -> None:
        self._specifications = specifications

    def is_satisfied_by(self, signals: ClassificationSignalsDTO) -> bool:
        return all(specification.is_satisfied_by(signals) for specification in self._specifications)
