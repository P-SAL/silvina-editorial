from unittest import TestCase

from src.domain.dtos.grammar_error_dto import GrammarErrorDTO
from src.domain.grammar.grammar_check_port import GrammarCheckPort
from src.domain.tests.grammar.fake_grammar_check_port import FakeGrammarCheckPort


class TestGrammarCheckPort(TestCase):
    def test_direct_instantiation_raises_type_error(self):
        with self.assertRaises(TypeError):
            GrammarCheckPort()

    def test_fake_check_returns_configured_error_list(self):
        error = GrammarErrorDTO(
            number=1,
            message="Test",
            context="ctx",
            offset=0,
            length=3,
            replacements=[],
        )
        fake = FakeGrammarCheckPort(errors=[error])
        result = fake.check(paragraphs=[])
        self.assertEqual(result, [error])

    def test_fake_check_returns_empty_list_when_no_errors_configured(self):
        fake = FakeGrammarCheckPort()
        result = fake.check(paragraphs=[])
        self.assertEqual(result, [])

    def test_fake_raises_configured_error(self):
        fake = FakeGrammarCheckPort(error=ValueError("boom"))
        with self.assertRaises(ValueError):
            fake.check(paragraphs=[])
