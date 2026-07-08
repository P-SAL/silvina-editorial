from src.domain.dtos.grammar_check_result_dto import GrammarCheckResultDTO
from src.domain.grammar.grammar_check_port import GrammarCheckPort
from src.domain.grammar.grammar_score_level import GrammarScoreLevel


class GrammarChecker:
    """Domain service that checks grammar and computes the resulting score level."""

    def __init__(self, grammar_check_port: GrammarCheckPort) -> None:
        self._grammar_check_port = grammar_check_port

    def check_grammar(self, paragraphs: list[str]) -> GrammarCheckResultDTO:
        """Check grammar errors in the given paragraphs and map them to a score level."""
        errors = self._grammar_check_port.check(paragraphs=paragraphs)
        level = GrammarScoreLevel.from_error_count(error_count=len(errors))
        return GrammarCheckResultDTO(score=level.score, feedback=level.feedback, errors=errors)
