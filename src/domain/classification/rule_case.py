from dataclasses import dataclass

from src.domain.classification.classification_specification import ClassificationSpecification
from src.domain.enums.article_type import ArticleType
from src.domain.enums.classification_confidence import ClassificationConfidence


@dataclass(frozen=True)
class RuleCase:
    """One row of the 17-case classification rule table."""

    specification: ClassificationSpecification
    article_type: ArticleType
    confidence: ClassificationConfidence | None
    reasoning_template: str
