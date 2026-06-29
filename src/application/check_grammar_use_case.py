from src.domain.dtos.grammar_check_result_dto import GrammarCheckResultDTO
from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler
from src.domain.grammar.grammar_check_port import GrammarCheckPort
from src.domain.grammar.grammar_score_level import GrammarScoreLevel


class CheckGrammarUseCase:
    def __init__(self, grammar_port: GrammarCheckPort) -> None:
        self._grammar_port = grammar_port

    @generic_error_handler
    def execute(self, paragraphs: list[str]) -> GrammarCheckResultDTO:
        errors = self._grammar_port.check(paragraphs=paragraphs)
        level = GrammarScoreLevel.from_error_count(error_count=len(errors))
        return GrammarCheckResultDTO(
            score=level.score,
            feedback=level.feedback,
            errors=errors,
        )
