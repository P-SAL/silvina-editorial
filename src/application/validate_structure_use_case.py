from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.dtos.structure_validation_result_dto import StructureValidationResultDTO
from src.domain.enums.article_type import ArticleType
from src.domain.enums.section_name import SectionName
from src.domain.exceptions.document_errors import DocumentEmpty
from src.domain.structure.structure_validator import StructureValidator


class ValidateStructureUseCase:
    """Application use case for document structure validation."""

    def __init__(self, validator: StructureValidator) -> None:
        self._validator = validator

    def execute(
        self,
        document_content: DocumentContentDTO,
        article_type: ArticleType,
        has_references: bool = False,
    ) -> StructureValidationResultDTO:
        if not document_content.paragraphs:
            raise DocumentEmpty

        _, missing = self._validator.validate(document_content, article_type)

        missing = [s for s in missing if s != SectionName.DEVELOPMENT]
        if has_references:
            missing = [s for s in missing if s != SectionName.REFERENCES]

        return StructureValidationResultDTO(
            is_valid=len(missing) == 0,
            missing_sections=list(missing),
        )
