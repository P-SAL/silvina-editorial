from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO


@dataclass(frozen=True)
class ClassificationSignalsDTO(BaseDTO):
    """Named replacement for legacy's positional s2a/s2b/s3/s4/s5/s6 tuple."""

    has_sufficient_reference_count: bool
    has_recent_references: bool
    has_methodological_vocabulary: bool
    has_research_intent: bool
    has_evidence_based_contribution: bool
    has_theoretical_justification: bool
