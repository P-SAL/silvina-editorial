from src.application.validate_apa_use_case import ValidateApaUseCase
from src.domain.citation.apa_validator import ApaValidator


class ValidateApaWiring:
    """Factory for building a ready-to-use ValidateApaUseCase."""

    def create_use_case(self) -> ValidateApaUseCase:
        return ValidateApaUseCase(validator=self._get_apa_validator())

    def _get_apa_validator(self) -> ApaValidator:
        return ApaValidator()
