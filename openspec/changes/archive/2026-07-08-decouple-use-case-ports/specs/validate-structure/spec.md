# Delta for validate-structure

## MODIFIED Requirements

### Requirement: StructureValidator Domain Service

`StructureValidator` MUST reside in `src/domain/structure/structure_validator.py`. It is enhanced to encapsulate the document empty guard and section post-filtering logic.
(Previously: Empty check and section post-filtering for DEVELOPMENT and REFERENCES were performed by the orchestrator.)

Method `validate_structure(document_content: DocumentContentDTO, article_type: ArticleType, has_references: bool) -> StructureValidationResultDTO` MUST perform:
1. If `document_content.paragraphs` is empty, raise `DocumentEmpty` exception.
2. Call `validate(document_content, article_type)` to get `(present, missing)` sections.
3. Filter the `missing` sections list:
   - Unconditionally remove `SectionName.DEVELOPMENT`.
   - Remove `SectionName.REFERENCES` if `has_references` is `True`.
4. Construct and return `StructureValidationResultDTO(is_valid=(len(filtered_missing) == 0), missing_sections=filtered_missing)`.

#### Scenario: Empty paragraphs list raises DocumentEmpty
- GIVEN a `DocumentContentDTO` with `paragraphs == []`
- WHEN `validate_structure` is called
- THEN it raises `DocumentEmpty` exception

#### Scenario: Post-filtering removes Development and conditionally removes References
- GIVEN a scientific article missing `DEVELOPMENT` and `REFERENCES` sections
- AND `has_references = True`
- WHEN `validate_structure` is called
- THEN both `DEVELOPMENT` and `REFERENCES` are omitted from the returned `missing_sections`
- AND `is_valid` is computed based on the filtered list

---

### Requirement: Orchestration — AnalyzeDocumentUseCase._validate_structure()

The private method `AnalyzeDocumentUseCase._validate_structure` MUST delegate directly to `StructureValidator.validate_structure` without local post-processing.
(Previously: Performed empty check and section filtering locally in the orchestrator before constructing the DTO.)

`_validate_structure(document_content: DocumentContentDTO, article_type: ArticleType, has_references: bool) -> StructureValidationResultDTO` MUST:
1. Call `self._structure_validator.validate_structure(document_content, article_type, has_references)`.
2. Return the result directly.

#### Scenario: Orchestration delegates structure validation
- GIVEN a valid `DocumentContentDTO`, `article_type`, and `has_references`
- WHEN `_validate_structure` is called
- THEN it returns the `StructureValidationResultDTO` produced by `StructureValidator.validate_structure`
