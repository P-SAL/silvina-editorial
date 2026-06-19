from dataclasses import dataclass, field

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.dimension_score_dto import DimensionScoreDTO
from src.domain.enums.quality_dimension import QualityDimension


@dataclass(frozen=True)
class ParsedResponseDTO(BaseDTO):
    """The result of parsing one LLM response into per-dimension scores."""

    scores: dict[QualityDimension, DimensionScoreDTO] = field(default_factory=dict)
    matched_dimensions: frozenset[QualityDimension] = field(default_factory=frozenset)
