from src.domain.citation.apa_validator import ApaValidator
from src.domain.dtos.apa_validation_result_dto import ApaValidationResultDTO


class ValidateApaUseCase:
    def __init__(self, validator: ApaValidator) -> None:
        self._validator = validator

    def execute(self, citations: list[tuple[str, int, str]]) -> ApaValidationResultDTO:
        if not citations:
            return ApaValidationResultDTO(is_valid=True, violation_count=0, violations=[])
        violations = self._validator.validate_all_citations(citations)
        count = len(violations)
        return ApaValidationResultDTO(
            is_valid=(count == 0), violation_count=count, violations=violations
        )
