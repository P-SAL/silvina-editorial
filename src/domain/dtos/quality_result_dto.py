from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.editorial_suitability_dto import EditorialSuitabilityDTO
from src.domain.enums.quality_level import QualityLevel


@dataclass(frozen=True)
class QualityResultDTO(BaseDTO):
    """Immutable result of quality analysis."""

    overall_score: float
    quality_level: QualityLevel
    dimension_scores: dict[str, dict[str, Any]] = field(default_factory=dict)
    timestamp: datetime = field(default_factory=datetime.now)
    editorial_suitability: EditorialSuitabilityDTO | None = None

    def __str__(self) -> str:
        """Return human-readable quality summary."""
        return f"Quality: {self.overall_score}/10 ({self.quality_level.value})"
