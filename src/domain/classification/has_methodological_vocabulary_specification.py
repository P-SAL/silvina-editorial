from src.domain.classification.classification_specification import ClassificationSpecification
from src.domain.dtos.classification_signals_dto import ClassificationSignalsDTO


class HasMethodologicalVocabularySpecification(ClassificationSpecification):
    """Satisfied when methodological vocabulary is detected (legacy signal S3)."""

    def is_satisfied_by(self, signals: ClassificationSignalsDTO) -> bool:
        return signals.has_methodological_vocabulary
