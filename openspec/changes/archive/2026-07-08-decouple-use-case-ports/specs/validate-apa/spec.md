# Delta for validate-apa

## MODIFIED Requirements

### Requirement: ApaValidator Domain Service — Behavioral Contract

`ApaValidator` MUST reside in `src/domain/citation/apa_validator.py` and remain stateless. Its method `validate_all_citations` is modified to encapsulate the filtering of `AUTHOR_YEAR` citations and paragraph text preview extraction internally.
(Previously: `validate_all_citations` accepted pre-filtered tuples `(citation_text, paragraph_index, paragraph_text)` from the orchestrator.)

`validate_all_citations(citations: list[CitationDTO], paragraphs: list[str]) -> list[ApaViolationDTO]` MUST perform:
1. Filter the input `citations` list to process only those with `citation_type == CitationType.AUTHOR_YEAR`.
2. For each filtered citation, retrieve the corresponding paragraph text from `paragraphs` using `citation.location`. If `citation.location` is out of bounds, fallback to an empty string.
3. Call `validate_citation(citation.text, citation.location, paragraph_text)` for each citation.
4. Accumulate and return the ordered list of all `ApaViolationDTO`s.

#### Scenario: Only AUTHOR_YEAR citations are validated and location preview constructed
- GIVEN a list containing one `AUTHOR_YEAR` citation at location 0 and one `NUMERIC` citation
- AND a list of paragraphs: `["Paragraph 0 text contents"]`
- WHEN `validate_all_citations` is called
- THEN only the `AUTHOR_YEAR` citation is validated
- AND the preview is created from `"Paragraph 0 text contents"`

#### Scenario: Empty citation list returns empty violations
- GIVEN an empty list of citations
- WHEN `validate_all_citations` is called
- THEN it returns an empty list `[]`

---

### Requirement: APA Validation Orchestration

The private method `AnalyzeDocumentUseCase._validate_apa` MUST accept `citations: list[CitationDTO]` and `paragraphs: list[str]` and delegate directly to `ApaValidator.validate_all_citations`.
(Previously: accepted pre-filtered tuples and performed empty-list guards inside the orchestrator.)

`_validate_apa(citations: list[CitationDTO], paragraphs: list[str]) -> ApaValidationResultDTO` MUST perform:
1. Call `apa_validator.validate_all_citations(citations=citations, paragraphs=paragraphs)`.
2. Compute `is_valid = (len(violations) == 0)` and `violation_count = len(violations)`.
3. Return `ApaValidationResultDTO(is_valid=is_valid, violation_count=violation_count, violations=violations)`.

#### Scenario: Orchestration computes is_valid and violation_count correctly
- GIVEN a list containing one citation with a conjunction error
- WHEN `_validate_apa` is called
- THEN it returns `ApaValidationResultDTO` with `is_valid=False` and `violation_count=1`
