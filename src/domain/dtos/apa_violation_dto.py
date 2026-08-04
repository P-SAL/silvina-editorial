from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO
from src.domain.enums.apa_error_type import ApaErrorType


@dataclass(frozen=True)
class ApaViolationDTO(BaseDTO):
    citation_text: str
    error_type: ApaErrorType
    location: int
    explanation: str
    correction: str
    paragraph_preview: str = ""
