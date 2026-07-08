# Delta for analyze-document

## ADDED Requirements

### Requirement: DocumentContentExtractor Domain Service

The `DocumentContentExtractor` domain service MUST reside in `src/domain/document/document_content_extractor.py`. It SHALL extract document content from a DOCX file using ports and handle count fallback logic.
- Constructor MUST inject: `document_text_port: DocumentTextPort`, `content_extraction_port: ContentExtractionPort`, and `character_count_port: CharacterCountPort`.
- Method `extract_content(docx_path: str) -> DocumentContentDTO`:
  1. Calls `document_text_port.read_paragraphs(path=docx_path)` to load paragraphs.
  2. Calls `content_extraction_port.extract(paragraphs, docx_path)` to get a base DTO.
  3. Calls `character_count_port.count(docx_path)`. On `CharacterCountUnavailable` or if result is `None`, returns the base DTO. Otherwise, returns a new DTO replacing word, char, and paragraph counts with the counted values.

#### Scenario: Content extraction executes successfully with count fallback
- GIVEN a valid DOCX path and `character_count_port` raises `CharacterCountUnavailable`
- WHEN `extract_content` is called
- THEN it returns a `DocumentContentDTO` containing text-based fallback counts

---

### Requirement: CitationExtractor Domain Service

The `CitationExtractor` domain service MUST reside in `src/domain/citation/citation_extractor.py`. It SHALL extract citations and references from a DOCX file.
- Constructor MUST inject: `citation_extraction_port: CitationExtractionPort` and `reference_extraction_port: ReferenceExtractionPort`.
- Method `extract_citations_and_references(docx_path: str) -> tuple[list[CitationDTO], list[ReferenceDTO], str]`:
  1. Calls `citation_extraction_port.extract_citations(docx_path=docx_path)`.
  2. Calls `reference_extraction_port.extract_references(docx_path=docx_path)`.
  3. Returns the tuple `(citations, references, section_type)`.

#### Scenario: Citations and references are extracted
- GIVEN a valid DOCX path
- WHEN `extract_citations_and_references` is called
- THEN it returns a tuple of citations, references, and references section type

---

### Requirement: DocumentFormatInspector Domain Service

The `DocumentFormatInspector` domain service MUST reside in `src/domain/document/document_format_inspector.py`. It SHALL inspect formatting rules.
- Constructor MUST inject: `document_format_inspection_port: DocumentFormatInspectionPort`.
- Method `inspect(docx_path: str, word_count: int) -> list[EumicViolationDTO]`:
  1. Calls `document_format_inspection_port.inspect(docx_path=docx_path, word_count=word_count)`.

#### Scenario: Format inspection finds violations
- GIVEN a document path and word count
- WHEN `inspect` is called
- THEN it returns a list of formatting violations

---

### Requirement: GrammarChecker Domain Service

The `GrammarChecker` domain service MUST reside in `src/domain/grammar/grammar_checker.py`. It SHALL perform grammar checks and compute score level.
- Constructor MUST inject: `grammar_check_port: GrammarCheckPort`.
- Method `check_grammar(paragraphs: list[str]) -> GrammarCheckResultDTO`:
  1. Calls `grammar_check_port.check(paragraphs=paragraphs)`.
  2. Maps error count using `GrammarScoreLevel.from_error_count(error_count=len(errors))`.
  3. Returns `GrammarCheckResultDTO(score=level.score, feedback=level.feedback, errors=errors)`.

#### Scenario: Grammar check returns errors and level
- GIVEN a list of paragraphs
- WHEN `check_grammar` is called
- THEN it returns a `GrammarCheckResultDTO` with grammar score and feedback

---

## MODIFIED Requirements

### Requirement: AnalyzeDocumentUseCase Orchestrator

`AnalyzeDocumentUseCase` MUST live in `src/application/analyze_document_use_case.py` and coordinate the document analysis steps. It accepts its 10 domain service dependencies via constructor injection:
- Domain services: `document_content_extractor`, `citation_extractor`, `document_format_inspector`, `grammar_checker`, `apa_validator`, `article_classifier`, `quality_analyzer`, `structure_validator`, `citation_matcher`, `recommendation_builder`.
(Previously: Accepted 7 ports, 5 domain services, and 1 builder — 13 dependencies total.)

Method `execute(document_path: str) -> ReportInputDTO` MUST be wrapped with `@generic_error_handler` and perform:
1. Extract content via `document_content_extractor.extract_content(document_path)`.
2. Extract citations/references via `citation_extractor.extract_citations_and_references(document_path)`.
3. Validate APA citations via `apa_validator.validate_all_citations(citations, document_content.paragraphs)`.
4. Grammar check via `grammar_checker.check_grammar(document_content.paragraphs)`.
5. Classify article via `article_classifier.classify(document_content)`.
6. Analyze quality via `quality_analyzer.analyze(document_content)`.
7. Validate structure via `structure_validator.validate_structure(document_content, classification.effective_structure_type, len(references) > 0)`.
8. Parse references section type to `SectionName`, falling back to `REFERENCES` on `ValueError`.
9. Match citations via `citation_matcher.match_citations_to_references(citations, references, section_name)`.
10. Verify format/EUMIC via `document_format_inspector.inspect(document_path, document_content.word_count)`.
11. Call `recommendation_builder.build(...)` -> `(recommendations, verdict)`.
12. Return `ReportInputDTO`.

#### Scenario: Orchestrator executes all pipeline steps sequentially
- GIVEN a valid `document_path`
- WHEN `execute(document_path)` is called
- THEN each of the 10 domain service dependencies is invoked and a `ReportInputDTO` is returned

#### Scenario: Structure validation uses effective structure type
- GIVEN a scientific article without "IMRyD" in reasoning
- WHEN structure validation is invoked
- THEN the orchestrator calls `structure_validator.validate_structure` with `article_type=ArticleType.DIVULGACION`

---

### Requirement: AnalyzeDocumentUseCaseWiring Assembly Factory

`AnalyzeDocumentUseCaseWiring` MUST reside in `src/infrastructure/wirings/analyze_document_use_case_wiring.py`. It MUST follow the private-method wiring pattern where:
- `create_use_case()` instantiates `EnvConfig` and delegates dependency injection to `_get_xxx()` private methods, injecting the configurations from the `EnvConfig` instance.
- Port helper methods instantiate and memoize/return infrastructure adapters.
- Domain service helper methods build the 10 domain services directly, wrapping ports or LLM generator dependencies as required.
- `_get_document_content_extractor()` returns `DocumentContentExtractor(self._get_document_text_port(), self._get_content_extraction_port(), self._get_character_count_port())`.
- `_get_citation_extractor()` returns `CitationExtractor(self._get_citation_extraction_port(), self._get_reference_extraction_port())`.
- `_get_document_format_inspector()` returns `DocumentFormatInspector(self._get_document_format_inspection_port())`.
- `_get_grammar_checker()` returns `GrammarChecker(self._get_grammar_check_port())`.
(Previously: Constructed and injected 7 ports and 5 domain services directly into the orchestrator.)

#### Scenario: Wiring constructs correct dependency graph
- GIVEN the wiring configuration
- WHEN `AnalyzeDocumentUseCaseWiring().create_use_case()` is called
- THEN it returns a valid `AnalyzeDocumentUseCase` with all 10 domain service dependencies injected

#### Scenario: Article classifier and quality analyzer share one LLM generator instance
- GIVEN `AnalyzeDocumentUseCaseWiring().create_use_case()`
- WHEN `result._article_classifier._llm_generator` and `result._quality_analyzer._llm_generator` are compared
- THEN they are the exact same object (`is`), not merely equal instances

#### Scenario: Environment variable overrides threshold at wiring time
- GIVEN `QUALITY_THRESHOLD=6.5` is set in the environment before `create_use_case()` instantiates `EnvConfig`
- WHEN `create_use_case()` is called
- THEN `recommendation_builder._settings.quality_threshold` equals `6.5`
