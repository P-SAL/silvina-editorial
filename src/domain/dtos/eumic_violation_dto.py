from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO
from src.domain.enums.severity_level import SeverityLevel


@dataclass(frozen=True)
class EumicViolationDTO(BaseDTO):
    """A single EUMIC editorial standard violation."""

    category: str
    message: str
    severity: SeverityLevel
    details: str = ""
