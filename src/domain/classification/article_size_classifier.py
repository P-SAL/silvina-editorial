from src.domain.enums.article_size import ArticleSize


class ArticleSizeClassifier:
    """Classifies an article's size based on its character count."""

    def classify(self, char_count: int) -> ArticleSize:
        """Classify article size based on character count with spaces."""
        if 36000 <= char_count <= 40000:
            return ArticleSize.LARGO
        if 16000 <= char_count <= 24000:
            return ArticleSize.CORTO
        if 24001 <= char_count <= 35999:
            return ArticleSize.NO_DEFINIDO
        return ArticleSize.FUERA_RANGO
