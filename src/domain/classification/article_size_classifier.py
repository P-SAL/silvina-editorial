from src.domain.enums.article_size import ArticleSize


class ArticleSizeClassifier:
    """Classifies an article's size based on its character count."""

    def __init__(
        self,
        *,
        short_min_chars: int = 16000,
        short_max_chars: int = 24000,
        undefined_min_chars: int = 24001,
        undefined_max_chars: int = 35999,
        long_min_chars: int = 36000,
        long_max_chars: int = 40000,
    ) -> None:
        self._short_min_chars = short_min_chars
        self._short_max_chars = short_max_chars
        self._undefined_min_chars = undefined_min_chars
        self._undefined_max_chars = undefined_max_chars
        self._long_min_chars = long_min_chars
        self._long_max_chars = long_max_chars

    def classify(self, char_count: int) -> ArticleSize:
        """Classify article size based on character count with spaces."""
        if self._long_min_chars <= char_count <= self._long_max_chars:
            return ArticleSize.LONG
        if self._short_min_chars <= char_count <= self._short_max_chars:
            return ArticleSize.SHORT
        if self._undefined_min_chars <= char_count <= self._undefined_max_chars:
            return ArticleSize.UNDEFINED
        return ArticleSize.OUT_OF_RANGE
