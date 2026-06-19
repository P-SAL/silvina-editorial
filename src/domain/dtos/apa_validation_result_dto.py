from dataclasses import dataclass, field

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.apa_violation_dto import ApaViolationDTO


@dataclass(frozen=True)
class ApaValidationResultDTO(BaseDTO):
    is_valid: bool
    violation_count: int
    violations: list[ApaViolationDTO] = field(default_factory=list)
