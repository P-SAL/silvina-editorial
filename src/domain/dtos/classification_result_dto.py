from dataclasses import dataclass, field
from datetime import datetime

from src.domain.dtos.base_dto import BaseDTO
from src.domain.enums.article_size import ArticleSize
from src.domain.enums.article_type import ArticleType


@dataclass(frozen=True)
class ClassificationResult(BaseDTO):
    """Immutable result of article classification."""

    article_type: ArticleType
    article_size: ArticleSize
    confidence: float | None
    reasoning: str
    timestamp: datetime = field(default_factory=datetime.now)

    @classmethod
    def create(
        cls,
        article_type: ArticleType,
        article_size: ArticleSize,
        confidence: float | None,
        reasoning: str,
    ) -> "ClassificationResult":
        """Build a ClassificationResult with all required fields."""
        return cls(
            article_type=article_type,
            article_size=article_size,
            confidence=confidence,
            reasoning=reasoning,
        )

    def __str__(self) -> str:
        """Return human-readable classification summary."""
        confidence_display = f"{self.confidence:.1%}" if self.confidence is not None else "—"
        return (
            f"Classification: {self.article_type.value} | "
            f"Size: {self.article_size.value} | "
            f"Confidence: {confidence_display}"
        )
