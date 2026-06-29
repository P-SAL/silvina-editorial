from unittest import TestCase

from src.application.check_grammar_use_case import CheckGrammarUseCase
from src.domain.dtos.grammar_check_result_dto import GrammarCheckResultDTO
from src.domain.dtos.grammar_error_dto import GrammarErrorDTO
from src.domain.exceptions.grammar_errors import GrammarCheckUnavailable
from src.domain.tests.grammar.fake_grammar_check_port import FakeGrammarCheckPort


def _make_errors(count: int) -> list[GrammarErrorDTO]:
    return [
        GrammarErrorDTO(
            number=index + 1,
            message="error",
            context="ctx",
            offset=0,
            length=4,
            replacements=[],
        )
        for index in range(count)
    ]


class TestCheckGrammarUseCase(TestCase):
    def test_execute_returns_perfect_score_when_no_errors(self):
        use_case = CheckGrammarUseCase(grammar_port=FakeGrammarCheckPort(errors=[]))
        result = use_case.execute(paragraphs=["text"])
        self.assertIsInstance(result, GrammarCheckResultDTO)
        self.assertEqual(result.score, 10.0)
        self.assertEqual(result.feedback, "Sin errores gramaticales")
        self.assertEqual(result.errors, [])

    def test_execute_returns_score_8_5_at_five_errors_boundary(self):
        use_case = CheckGrammarUseCase(grammar_port=FakeGrammarCheckPort(errors=_make_errors(5)))
        result = use_case.execute(paragraphs=["text"])
        self.assertEqual(result.score, 8.5)
        self.assertEqual(result.feedback, "Pocos errores gramaticales")

    def test_execute_returns_score_7_0_at_fifteen_errors_boundary(self):
        use_case = CheckGrammarUseCase(grammar_port=FakeGrammarCheckPort(errors=_make_errors(15)))
        result = use_case.execute(paragraphs=["text"])
        self.assertEqual(result.score, 7.0)
        self.assertEqual(result.feedback, "Errores gramaticales moderados")

    def test_execute_returns_score_5_0_at_sixteen_errors_threshold(self):
        use_case = CheckGrammarUseCase(grammar_port=FakeGrammarCheckPort(errors=_make_errors(16)))
        result = use_case.execute(paragraphs=["text"])
        self.assertEqual(result.score, 5.0)
        self.assertEqual(result.feedback, "Muchos errores gramaticales")

    def test_execute_propagates_grammar_check_unavailable_from_port(self):
        use_case = CheckGrammarUseCase(
            grammar_port=FakeGrammarCheckPort(error=GrammarCheckUnavailable()),
        )
        with self.assertRaises(GrammarCheckUnavailable):
            use_case.execute(paragraphs=["text"])
