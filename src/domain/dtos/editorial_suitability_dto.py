from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO


@dataclass(frozen=True)
class EditorialSuitabilityDTO(BaseDTO):
    """Immutable qualitative editorial suitability evaluation."""

    contribution_verdict: str
    contribution_phrase: str
    contribution_observation: str
    alignment_verdict: str
    alignment_lines: str
    alignment_justification: str
