from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO
from src.domain.enums.publication_verdict import PublicationVerdict


@dataclass(frozen=True)
class PublicationVerdictDTO(BaseDTO):
    """Immutable final publication verdict for an analyzed document."""

    verdict: PublicationVerdict
    message: str
