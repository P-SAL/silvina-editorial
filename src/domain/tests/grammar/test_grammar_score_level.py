from unittest import TestCase

from src.domain.grammar.grammar_score_level import GrammarScoreLevel


class TestGrammarScoreLevel(TestCase):
    def test_returns_perfect_for_zero_errors(self):
        level = GrammarScoreLevel.from_error_count(error_count=0)
        self.assertIs(level, GrammarScoreLevel.PERFECT)
        self.assertEqual(level.score, 10.0)
        self.assertEqual(level.feedback, "Sin errores gramaticales")

    def test_returns_minor_for_one_to_five_errors(self):
        for count in (1, 5):
            with self.subTest(count=count):
                level = GrammarScoreLevel.from_error_count(error_count=count)
                self.assertIs(level, GrammarScoreLevel.MINOR)
                self.assertEqual(level.score, 8.5)
                self.assertEqual(level.feedback, "Pocos errores gramaticales")

    def test_returns_moderate_for_six_to_fifteen_errors(self):
        for count in (6, 15):
            with self.subTest(count=count):
                level = GrammarScoreLevel.from_error_count(error_count=count)
                self.assertIs(level, GrammarScoreLevel.MODERATE)
                self.assertEqual(level.score, 7.0)
                self.assertEqual(level.feedback, "Errores gramaticales moderados")

    def test_returns_severe_for_sixteen_or_more_errors(self):
        for count in (16, 100):
            with self.subTest(count=count):
                level = GrammarScoreLevel.from_error_count(error_count=count)
                self.assertIs(level, GrammarScoreLevel.SEVERE)
                self.assertEqual(level.score, 5.0)
                self.assertEqual(level.feedback, "Muchos errores gramaticales")
