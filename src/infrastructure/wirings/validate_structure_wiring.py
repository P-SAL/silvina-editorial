from src.application.validate_structure_use_case import ValidateStructureUseCase
from src.domain.structure.structure_validator import StructureValidator


class ValidateStructureWiring:
    """Factory for building a ready-to-use ValidateStructureUseCase."""

    def create_use_case(self) -> ValidateStructureUseCase:
        return ValidateStructureUseCase(validator=self._get_structure_validator())

    def _get_structure_validator(self) -> StructureValidator:
        return StructureValidator()
