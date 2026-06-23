from src.domain.classification.classification_specification import ClassificationSpecification
from src.domain.dtos.classification_signals_dto import ClassificationSignalsDTO


class HasEvidenceBasedContributionSpecification(ClassificationSpecification):
    """Satisfied when an evidence-based contribution is detected (legacy signal S5)."""

    def is_satisfied_by(self, signals: ClassificationSignalsDTO) -> bool:
        return signals.has_evidence_based_contribution
