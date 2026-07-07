# Specification: Document Analysis Orchestrator & Recommendations (Slice 13 + Refactor)

> Normative guide: `.agent/skills/clean-architecture/SKILL.md`
> Parent proposal: `openspec/changes/analyze-document-orchestrator/proposal.md`
> Migration plan reference: `docs/plan-migracion-hexagonal.md`
> **Refactor note**: Slice 14 integration observations caused intentional divergences from the original Slice 13 spec. This document reflects the final implemented state.

---

## Purpose

This specification defines the behavior of the document analysis pipeline orchestrator, the domain service that generates editorial recommendations, and the supporting DTOs and configuration settings in compliance with clean hexagonal architecture guidelines.

---

## Requirements

### Requirement: RecommendationPriority Enum (3 members)

The `RecommendationPriority` enum MUST live in `src/domain/enums/recommendation_priority.py` and carry exactly **three** values representing severity levels for specific editorial recommendations:

- `HIGH` maps to `"alta"`
- `MEDIUM` maps to `"media"`
- `LOW` maps to `"baja"`

> **Rationale**: `CRITICAL`, `WARNING`, and `APPROVED` were removed from this enum because they represent a *publication verdict* — a semantically distinct concept. They live in `PublicationVerdict` (see below).

#### Scenario: RecommendationPriority contains exactly three members

- GIVEN the `RecommendationPriority` enum
- WHEN the list of members is queried
- THEN it has exactly 3 members: `HIGH`, `MEDIUM`, and `LOW`

#### Scenario: RecommendationPriority values match Spanish strings

- GIVEN the `RecommendationPriority` members
- WHEN their values are read
- THEN the values map to `"alta"`, `"media"`, and `"baja"` respectively

---

### Requirement: PublicationVerdict Enum

`PublicationVerdict` MUST live in `src/domain/enums/publication_verdict.py` and carry exactly **three** values representing the final editorial verdict for a document:

- `CRITICAL` maps to `"critica"`
- `WARNING` maps to `"advertencia"`
- `APPROVED` maps to `"aprobado"`

#### Scenario: PublicationVerdict contains exactly three members

- GIVEN the `PublicationVerdict` enum
- WHEN the list of members is queried
- THEN it has exactly 3 members: `CRITICAL`, `WARNING`, and `APPROVED`

---

### Requirement: PublicationVerdictDTO

`PublicationVerdictDTO` MUST be a frozen data transfer object inheriting from `BaseDTO` in `src/domain/dtos/publication_verdict_dto.py` with fields:
- `verdict: PublicationVerdict`
- `message: str`

#### Scenario: PublicationVerdictDTO is immutable

- GIVEN a constructed `PublicationVerdictDTO`
- WHEN any field is modified or reassigned
- THEN Python raises `FrozenInstanceError`

---

### Requirement: Immutable RecommendationDTO

`RecommendationDTO` MUST be a frozen data transfer object inheriting from `BaseDTO` in `src/domain/dtos/recommendation_dto.py` with fields:
- `priority: RecommendationPriority`
- `message: str`

#### Scenario: RecommendationDTO is immutable

- GIVEN a constructed `RecommendationDTO`
- WHEN any field is modified or reassigned
- THEN Python raises `FrozenInstanceError`

---

### Requirement: ClassificationResultDTO IMRyD Override Property

`ClassificationResultDTO` (in `src/domain/dtos/classification_result_dto.py`) MUST expose a read-only property `effective_structure_type: ArticleType` using guard clauses:

1. If `article_type` is NOT `ArticleType.CIENTIFICO`, return `article_type` immediately.
2. If `"IMRyD"` is in `(self.reasoning or "")` (case-sensitive), return `ArticleType.CIENTIFICO`.
3. Otherwise return `ArticleType.DIVULGACION`.

#### Scenario: Scientific article with IMRyD reasoning

- GIVEN a `ClassificationResultDTO` with `article_type = ArticleType.CIENTIFICO` and `reasoning` containing `"IMRyD"`
- WHEN `effective_structure_type` is called
- THEN it returns `ArticleType.CIENTIFICO`

#### Scenario: Scientific article without IMRyD reasoning

- GIVEN a `ClassificationResultDTO` with `article_type = ArticleType.CIENTIFICO` and `reasoning` not containing `"IMRyD"`
- WHEN `effective_structure_type` is called
- THEN it returns `ArticleType.DIVULGACION`

#### Scenario: Non-scientific article type

- GIVEN a `ClassificationResultDTO` with `article_type = ArticleType.DIVULGACION`
- WHEN `effective_structure_type` is called
- THEN it returns `ArticleType.DIVULGACION` unchanged

---

### Requirement: RecommendationSettingsDTO DTO (no defaults)

`RecommendationSettingsDTO` MUST be a frozen DTO inheriting from `BaseDTO` in `src/domain/dtos/recommendation_settings_dto.py`. All fields are **required** (no default values). Default values are the responsibility of the infrastructure config layer (`EnvConfig`):

- `publish_threshold: float`
- `quality_threshold: float`
- `grammar_threshold: float`
- `dimension_threshold: float`
- `citation_match_threshold: float`
- `critical_citation_match_threshold: float`
- `citation_count_threshold: int`
- `classification_confidence_threshold: float`
- `critical_quality_threshold: float`
- `critical_grammar_threshold: float`

> **Rationale**: Defaults in domain DTOs introduce hidden dependencies on infrastructure decisions. Defaults belong in `EnvConfig` (infrastructure config layer) and are resolved at wiring time.

---

### Requirement: EnvConfig Infrastructure Config Class

`EnvConfig` MUST reside in `src/infrastructure/env_config.py`. It MUST parse environment variables at instantiation, cast them, and cache them as typed instance attributes. It MUST expose a method `get_recommendation_settings() -> RecommendationSettingsDTO` to build recommendation settings.

| Env var | Type | Default | Attribute |
|---|---|---|---|
| `CITATION_MAX_AUTHOR_NAME_LENGTH` | `int` | `100` | `citation_max_author_name_length` |
| `GRAMMAR_MAX_REPLACEMENTS` | `int` | `3` | `grammar_max_replacements` |
| `STRUCTURE_MAX_HEADER_LENGTH` | `int` | `100` | `structure_max_header_length` |
| `ARTICLE_CLASSIFIER_TEMPERATURE` | `float` | `0.1` | `article_classifier_temperature` |
| `ARTICLE_CLASSIFIER_NUM_PREDICT` | `int` | `300` | `article_classifier_num_predict` |
| `ARTICLE_SIZE_SHORT_MIN_CHARS` | `int` | `16000` | `article_size_short_min_chars` |
| `ARTICLE_SIZE_SHORT_MAX_CHARS` | `int` | `24000` | `article_size_short_max_chars` |
| `ARTICLE_SIZE_UNDEFINED_MIN_CHARS` | `int` | `24001` | `article_size_undefined_min_chars` |
| `ARTICLE_SIZE_UNDEFINED_MAX_CHARS` | `int` | `35999` | `article_size_undefined_max_chars` |
| `ARTICLE_SIZE_LONG_MIN_CHARS` | `int` | `36000` | `article_size_long_min_chars` |
| `ARTICLE_SIZE_LONG_MAX_CHARS` | `int` | `40000` | `article_size_long_max_chars` |
| `QUALITY_LEVEL_EXCELLENT_THRESHOLD` | `float` | `9.0` | `quality_level_excellent_threshold` |
| `QUALITY_LEVEL_GOOD_THRESHOLD` | `float` | `7.0` | `quality_level_good_threshold` |
| `QUALITY_LEVEL_ACCEPTABLE_THRESHOLD` | `float` | `5.0` | `quality_level_acceptable_threshold` |
| `QUALITY_LEVEL_NEEDS_IMPROVEMENT_THRESHOLD` | `float` | `3.0` | `quality_level_needs_improvement_threshold` |
| `QUALITY_MIN_SAMPLE_WORD_COUNT` | `int` | `400` | `quality_min_sample_word_count` |
| `QUALITY_TEXT_SAMPLE_CHARACTER_LIMIT` | `int` | `8000` | `quality_text_sample_character_limit` |
| `OLLAMA_MODEL_NAME` | `str` | `"llama3-gradient:8b-instruct-1048k-q4_K_M"` | `ollama_model_name` |
| `OLLAMA_BASE_URL` | `str` | `"http://localhost:11434"` | `ollama_base_url` |
| `PUBLISH_THRESHOLD` | `float` | `7.0` | `publish_threshold` |
| `QUALITY_THRESHOLD` | `float` | `7.0` | `quality_threshold` |
| `GRAMMAR_THRESHOLD` | `float` | `7.0` | `grammar_threshold` |
| `DIMENSION_THRESHOLD` | `float` | `6.0` | `dimension_threshold` |
| `CITATION_MATCH_THRESHOLD` | `float` | `90.0` | `citation_match_threshold` |
| `CRITICAL_CITATION_MATCH_THRESHOLD` | `float` | `50.0` | `critical_citation_match_threshold` |
| `CITATION_COUNT_THRESHOLD` | `int` | `10` | `citation_count_threshold` |
| `CLASSIFICATION_CONFIDENCE_THRESHOLD` | `float` | `0.7` | `classification_confidence_threshold` |
| `CRITICAL_QUALITY_THRESHOLD` | `float` | `5.0` | `critical_quality_threshold` |
| `CRITICAL_GRAMMAR_THRESHOLD` | `float` | `5.0` | `critical_grammar_threshold` |
| `SILVINA_APP_NAME` | `str` | `"Silvina Editorial Assistant"` | `silvina_app_name` |
| `SILVINA_VERSION` | `str` | `"0.9"` | `silvina_version` |
| `REPORT_SCORE_HIGH_THRESHOLD` | `float` | `8.0` | `report_score_high_threshold` |
| `REPORT_SCORE_MEDIUM_THRESHOLD` | `float` | `6.0` | `report_score_medium_threshold` |
| `REPORT_WORDS_PER_PAGE` | `int` | `250` | `report_words_per_page` |
| `REPORT_MAX_ERRORS_DISPLAYED` | `int` | `5` | `report_max_errors_displayed` |
| `REPORT_CONTEXT_TRUNCATION_LIMIT` | `int` | `150` | `report_context_truncation_limit` |
| `REPORT_MAX_REPLACEMENTS` | `int` | `3` | `report_max_replacements` |

> **Naming note**: the `PUBLISH_THRESHOLD` … `CRITICAL_GRAMMAR_THRESHOLD` variables (recommendation thresholds) carry no `RECOMMENDATION_` prefix. In `.env`/`.env.example` they MUST be grouped under a section comment (e.g. `# Recommendation thresholds`) instead of relying on a name prefix for grouping.

#### Scenario: EnvConfig defaults are loaded when env is empty

- GIVEN an empty environment
- WHEN `EnvConfig` is instantiated
- THEN attributes match the defaults in the table above

#### Scenario: EnvConfig parses and casts environment variables

- GIVEN environment contains `CITATION_MAX_AUTHOR_NAME_LENGTH=150`
- WHEN `EnvConfig` is instantiated
- THEN `env_config.citation_max_author_name_length == 150`

---

### Requirement: AnalysisContext Value Object

`AnalysisContext` MUST reside in `src/domain/recommendation/analysis_context.py` as a frozen dataclass grouping all analysis inputs and settings required by recommendation rules:

- `classification: ClassificationResultDTO`
- `quality: QualityResultDTO`
- `structure: StructureValidationResultDTO`
- `citations: CitationAnalysisResultDTO`
- `apa_validation: ApaValidationResultDTO`
- `grammar: GrammarCheckResultDTO`
- `settings: RecommendationSettingsDTO`

It MUST expose a computed property `citation_match_rate: float` returning `100.0` when `total_citations == 0`, otherwise `matched_count / total_citations * 100.0`.

---

### Requirement: RecommendationRule Abstract Base

`RecommendationRule` MUST reside in `src/domain/recommendation/recommendation_rule.py` as an abstract base class with a single abstract method:

```python
def evaluate(self, context: AnalysisContext) -> list[RecommendationDTO]: ...
```

---

### Requirement: Concrete Recommendation Rules

Seven concrete rule classes MUST reside in `src/domain/recommendation/`, one class per file, each implementing `RecommendationRule`:

1. **`QualityRule`** (`quality_rule.py`) — `HIGH` when `quality.overall_score < settings.quality_threshold`
2. **`GrammarRule`** (`grammar_rule.py`) — `HIGH` when `grammar.score < settings.grammar_threshold`
3. **`DimensionRule`** (`dimension_rule.py`) — `MEDIUM` for each dimension score below `settings.dimension_threshold`
4. **`StructureRule`** (`structure_rule.py`) — `HIGH` for each missing section when `structure.is_valid is False`
5. **`CitationMatchRule`** (`citation_match_rule.py`) — `HIGH` when match rate below `citation_match_threshold`; `MEDIUM` when unmatched count > 0 but above threshold
6. **`CitationCountRule`** (`citation_count_rule.py`) — `MEDIUM` when `total_citations < settings.citation_count_threshold`
7. **`ConfidenceRule`** (`confidence_rule.py`) — `LOW` when `classification.confidence` is not `None` and below `settings.classification_confidence_threshold`

---

### Requirement: PublicationVerdictEvaluator Domain Service

`PublicationVerdictEvaluator` MUST reside in `src/domain/recommendation/publication_verdict_evaluator.py`. Its method `evaluate(context: AnalysisContext) -> PublicationVerdictDTO` determines the final publication verdict:

- **`CRITICAL`** if any of: `quality < 5.0`, `grammar < 5.0`, `structure.is_valid is False`, `match_rate < critical_citation_match_threshold`, OR `total_citations == 0`
- **`WARNING`** if any of: `quality < publish_threshold`, `grammar < publish_threshold`, `match_rate < citation_match_threshold`, `len(apa_validation.violations) > 0`
- **`APPROVED`** otherwise

---

### Requirement: RecommendationBuilder Domain Service (Rule Pattern)

`RecommendationBuilder` MUST reside in `src/domain/recommendation/recommendation_builder.py`. Its constructor accepts:
- `settings: RecommendationSettingsDTO`
- `rules: list[RecommendationRule] | None` (defaults to the 7 concrete rules)
- `verdict_evaluator: PublicationVerdictEvaluator | None` (defaults to a new instance)

Its method `build(...) -> tuple[list[RecommendationDTO], PublicationVerdictDTO]` MUST:
1. Build an `AnalysisContext` from the inputs and settings.
2. Call `rule.evaluate(context)` for each rule and concatenate results.
3. Call `verdict_evaluator.evaluate(context)` to get the verdict.
4. Return `(recommendations, verdict)`.

> **Rationale**: The Rule pattern makes each recommendation rule independently testable and eliminates long `if/elif` chains. The verdict is a distinct return value (not a recommendation) to reflect its semantic difference.

#### Scenario: All thresholds satisfied — APPROVED verdict, empty recommendations

- GIVEN `RecommendationBuilder` with default rules and all analysis inputs satisfying thresholds
- WHEN `build()` is executed
- THEN it returns `([], PublicationVerdictDTO(verdict=PublicationVerdict.APPROVED, ...))`

#### Scenario: Quality below threshold — HIGH recommendation + WARNING verdict

- GIVEN a quality score of `6.5` (above critical but below `quality_threshold=7.0`)
- WHEN `build()` is executed
- THEN the list contains a `HIGH` recommendation for quality AND the verdict is `WARNING`

#### Scenario: Critical quality issue — CRITICAL verdict

- GIVEN a quality score of `4.5` (below critical threshold `5.0`)
- WHEN `build()` is executed
- THEN the verdict is `CRITICAL`

#### Scenario: Zero citations — CRITICAL verdict

- GIVEN `total_citations == 0`
- WHEN `build()` is executed
- THEN the verdict is `CRITICAL` with a message about missing APA citations

---

### Requirement: AnalyzeDocumentUseCase Orchestrator

> **Modified (2026-07-04, `refactor_analyze_document_wiring`)**: The 10 pass-through
> sub-use cases (`ReadDocumentUseCase`, `ExtractContentUseCase`, `ExtractCitationsUseCase`,
> `ValidateApaUseCase`, `CheckGrammarUseCase`, `ClassifyArticleUseCase`, `AnalyzeQualityUseCase`,
> `ValidateStructureUseCase`, `MatchCitationsUseCase`, `VerifyEumicUseCase`) and their 10
> wirings were deleted. `AnalyzeDocumentUseCase` now orchestrates 7 ports, 5 domain services,
> and 1 builder directly — 13 dependencies total (not 11). Fallback/filtering logic that used
> to live inside the deleted sub-use cases (`CharacterCountUnavailable` handling, APA's
> empty-citations guard, grammar scoring, structure's `DEVELOPMENT`/`REFERENCES` filtering) was
> moved into private methods on `AnalyzeDocumentUseCase` itself (`_extract_content`,
> `_validate_apa`, `_check_grammar`, `_validate_structure`). See each capability's own spec
> (`read-document`, `extract-content`, `extract-citations`, `validate-apa`, `check-grammar`,
> `classify-article`, `analyze-quality`, `validate-structure`, `validate-citations`) for the
> "Superseded" note pointing back here.

`AnalyzeDocumentUseCase` MUST live in `src/application/analyze_document_use_case.py` and
coordinate the document analysis steps. It accepts its 13 dependencies via constructor
injection:

- Ports: `document_text_port`, `content_extraction_port`, `character_count_port`,
  `citation_extraction_port`, `reference_extraction_port`, `grammar_check_port`,
  `document_format_inspection_port`
- Domain services: `apa_validator`, `article_classifier`, `quality_analyzer`,
  `structure_validator`, `citation_matcher`
- Builder: `recommendation_builder`

Method `execute(document_path: str) -> ReportInputDTO` MUST be wrapped with `@generic_error_handler`
(the single error-handling layer for the whole pipeline — none of its collaborator ports/services
carry their own `@generic_error_handler` at the orchestration level) and perform:

1. Load paragraphs via `document_text_port.read_paragraphs(path=document_path)`.
2. Extract content via private method `_extract_content(paragraphs, docx_path=document_path)`,
   which calls `content_extraction_port.extract()` then merges accurate counts from
   `character_count_port.count()`, falling back to text-based counts on `CharacterCountUnavailable`
   or a `None` result.
3. Extract citations via `citation_extraction_port.extract_citations(docx_path=document_path)`
   and references via `reference_extraction_port.extract_references(docx_path=document_path)`.
4. Filter `AUTHOR_YEAR` citations and pass tuples `(text, location, paragraph_text)` to private
   method `_validate_apa(citations)` (empty-citations guard: returns
   `ApaValidationResultDTO(is_valid=True, violation_count=0, violations=[])` immediately without
   calling `apa_validator` when the filtered list is empty).
5. Grammar check via private method `_check_grammar(paragraphs)`, which calls
   `grammar_check_port.check()` and computes score/feedback from the error count via
   `GrammarScoreLevel.from_error_count()`.
6. Classify article via `article_classifier.classify(document_content=document_content)`.
7. Analyze quality via `quality_analyzer.analyze(document_content=..., article_type=classification.article_type)`.
8. Validate structure via private method `_validate_structure(document_content, article_type=classification.effective_structure_type, has_references=len(references) > 0)`
   (raises `DocumentEmpty` if `document_content.paragraphs` is empty; removes `SectionName.DEVELOPMENT`
   unconditionally and `SectionName.REFERENCES` when `has_references` is `True`).
9. Parse `section_type` to `SectionName`; fallback to `SectionName.REFERENCES` on `ValueError`.
10. Match citations via `citation_matcher.match_citations_to_references(citations=..., references=..., section_type=section_name)`.
11. Verify EUMIC via `document_format_inspection_port.inspect(docx_path=document_path, word_count=document_content.word_count)` (no fatal exception on violations).
12. Call `recommendation_builder.build()` and **unpack the tuple**: `recommendations, verdict = builder.build(...)`.
13. Construct and return `ReportInputDTO` with all results including `verdict`.

#### Scenario: Orchestrator executes all pipeline steps sequentially

- GIVEN a valid `document_path`
- WHEN `execute(document_path)` is called
- THEN each port/domain-service collaborator is called exactly once and a `ReportInputDTO` is returned

#### Scenario: Only AUTHOR_YEAR citations are validated

- GIVEN a document with 2 `AUTHOR_YEAR` and 1 `NUMERIC` citations
- WHEN the orchestrator runs
- THEN only the 2 `AUTHOR_YEAR` citations are sent to `_validate_apa` (and, since non-empty,
  forwarded to `apa_validator.validate_all_citations()`)

#### Scenario: Structure validation uses effective structure type

- GIVEN a `CIENTIFICO` article without `"IMRyD"` in reasoning
- WHEN structure validation is invoked
- THEN `structure_validator.validate()` receives `article_type=ArticleType.DIVULGACION`

#### Scenario: has_references correctly filters REFERENCES from missing sections

- GIVEN `structure_validator.validate()` returns `SectionName.REFERENCES` among the missing sections
- WHEN `_validate_structure` is called with `has_references=True`
- THEN `SectionName.REFERENCES` is NOT in the resulting `missing_sections`
- WHEN `_validate_structure` is called with `has_references=False`
- THEN `SectionName.REFERENCES` IS in the resulting `missing_sections`

---

### Requirement: AnalyzeDocumentUseCaseWiring Assembly Factory

> **Modified (2026-07-04, `refactor_analyze_document_wiring`)**: `AnalyzeDocumentUseCaseWiring`
> is now the sole composition root for the analysis pipeline — it no longer delegates to 10
> sub-wirings. Each `_get_xxx()` method builds its adapter or domain service directly.

`AnalyzeDocumentUseCaseWiring` MUST reside in `src/infrastructure/wirings/analyze_document_use_case_wiring.py`.
It MUST follow the **private-method wiring pattern**:

- `create_use_case()` instantiates `EnvConfig` and delegates dependency injection to `_get_xxx()` private methods, injecting the configurations from the `EnvConfig` instance.
- Each `_get_xxx()` method instantiates the corresponding adapter or domain service directly (no sub-wiring classes remain).
- `_get_document_text_port()` returns `DocxTextAdapter()`.
- `_get_content_extraction_port()` returns `ParagraphContentAdapter()`.
- `_get_character_count_port()` returns `Win32ComWordCountAdapter()`.
- `_get_citation_extraction_port()` returns `DocxCitationAdapter(document_text_port=self._get_document_text_port(), max_author_name_length=env_config.citation_max_author_name_length)`.
- `_get_reference_extraction_port()` returns `DocxReferenceAdapter(document_text_port=self._get_document_text_port())`.
- `_get_grammar_check_port()` returns `LanguageToolAdapter(max_replacements=env_config.grammar_max_replacements)`.
- `_get_document_format_inspection_port()` returns `DocxEumicAdapter()`.
- `_get_apa_validator()` returns `ApaValidator()`.
- `_get_citation_matcher()` returns `CitationMatcher()`.
- `_get_structure_validator()` returns `StructureValidator(max_header_length=env_config.structure_max_header_length)`.
- `_get_article_classifier()` and `_get_quality_analyzer()` each construct their domain service directly, both consuming the **same shared** `LlmGeneratorPort` instance from `_get_llm_generator()` (memoized on the wiring instance) rather than each assembling their own `OllamaGeneratorAdapter`. These services are constructed with configuration values injected from `EnvConfig`.
- `_get_recommendation_builder()` obtains settings from `EnvConfig.get_recommendation_settings()`.

#### Scenario: Wiring constructs correct dependency graph

- GIVEN the wiring configuration
- WHEN `AnalyzeDocumentUseCaseWiring().create_use_case()` is called
- THEN it returns a valid `AnalyzeDocumentUseCase` with all 13 dependencies injected

#### Scenario: Article classifier and quality analyzer share one LLM generator instance

- GIVEN `AnalyzeDocumentUseCaseWiring().create_use_case()`
- WHEN `result._article_classifier._llm_generator` and `result._quality_analyzer._llm_generator` are compared
- THEN they are the exact same object (`is`), not merely equal instances

#### Scenario: Environment variable overrides threshold at wiring time

- GIVEN `QUALITY_THRESHOLD=6.5` is set in the environment before `create_use_case()` instantiates `EnvConfig`
- WHEN `create_use_case()` is called
- THEN `recommendation_builder._settings.quality_threshold` equals `6.5`
