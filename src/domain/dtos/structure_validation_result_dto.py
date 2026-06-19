from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

from src.domain.dtos.base_dto import BaseDTO


@dataclass(frozen=True)
class StructureValidationResultDTO(BaseDTO):
    """Immutable result of structure validation."""

    is_valid: bool
    missing_sections: list[str] = field(default_factory=list)
    section_details: dict[str, dict[str, Any]] = field(default_factory=dict)
    timestamp: datetime = field(default_factory=datetime.now)

    def __str__(self) -> str:
        """Return human-readable structure validation summary."""
        if self.is_valid:
            return "Structure: Valid"
        return f"Structure: Invalid ({len(self.missing_sections)} missing)"
