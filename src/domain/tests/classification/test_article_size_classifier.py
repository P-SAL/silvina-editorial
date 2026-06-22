from unittest import TestCase

from src.domain.classification.article_size_classifier import ArticleSizeClassifier
from src.domain.enums.article_size import ArticleSize


class TestArticleSizeClassifier(TestCase):
    def test_each_threshold_boundary_maps_to_correct_article_size(self):
        classifier = ArticleSizeClassifier()

        self.assertEqual(classifier.classify(16000), ArticleSize.CORTO)
        self.assertEqual(classifier.classify(24000), ArticleSize.CORTO)
        self.assertEqual(classifier.classify(24001), ArticleSize.NO_DEFINIDO)
        self.assertEqual(classifier.classify(35999), ArticleSize.NO_DEFINIDO)
        self.assertEqual(classifier.classify(36000), ArticleSize.LARGO)
        self.assertEqual(classifier.classify(40000), ArticleSize.LARGO)
        self.assertEqual(classifier.classify(40001), ArticleSize.FUERA_RANGO)
