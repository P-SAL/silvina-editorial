from unittest import TestCase

from src.domain.classification.article_classification_response_parser import (
    ArticleClassificationResponseParser,
)


class TestArticleClassificationResponseParser(TestCase):
    def test_well_formed_response_parses_all_three_signals_correctly(self):
        response_text = "S4: SI\nS5: NO\nS6: SI"
        parser = ArticleClassificationResponseParser()

        result = parser.parse(response_text)

        self.assertEqual(result, (True, False, True))

    def test_malformed_response_yields_all_false_without_raising(self):
        response_text = "respuesta sin los marcadores esperados"
        parser = ArticleClassificationResponseParser()

        result = parser.parse(response_text)

        self.assertEqual(result, (False, False, False))
