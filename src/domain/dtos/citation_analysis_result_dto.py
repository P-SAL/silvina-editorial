from dataclasses import dataclass, field
from datetime import datetime

from src.domain.dtos.base_dto import BaseDTO


@dataclass(frozen=True)
class CitationAnalysisResult(BaseDTO):
    """Immutable result of citation analysis."""

    total_citations: int
    total_references: int
    matched_count: int
    unmatched_count: int
    citations_by_type: dict[str, int] = field(default_factory=dict)
    unmatched_citations: list[str] = field(default_factory=list)
    timestamp: datetime = field(default_factory=datetime.now)

    def __str__(self) -> str:
        """Return human-readable citation analysis summary."""
        if self.total_citations == 0:
            return f"Citations: {self.total_citations} (0.0% matched)"
        match_rate = self.matched_count / self.total_citations * 100
        return f"Citations: {self.total_citations} ({match_rate:.1f}% matched)"
