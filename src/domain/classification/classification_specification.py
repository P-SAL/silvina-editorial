from abc import ABC, abstractmethod

from src.domain.dtos.classification_signals_dto import ClassificationSignalsDTO


class ClassificationSpecification(ABC):
    """Predicate over classification signals, used as a rule-table row condition."""

    @abstractmethod
    def is_satisfied_by(self, signals: ClassificationSignalsDTO) -> bool:
        """Return True if the given signals satisfy this specification."""
