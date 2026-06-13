from unittest import TestCase

from src.domain.enums.article_size import ArticleSize


class TestArticleSize(TestCase):
    def test_members_and_values(self):
        self.assertEqual(ArticleSize.LARGO.value, "largo")
        self.assertEqual(ArticleSize.CORTO.value, "corto")
        self.assertEqual(ArticleSize.NO_DEFINIDO.value, "no_definido")
        self.assertEqual(ArticleSize.FUERA_RANGO.value, "fuera_rango")

    def test_member_count(self):
        self.assertEqual(len(ArticleSize), 4)
