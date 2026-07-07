# Delta for Document Analysis (centralize-configuration)

## ADDED Requirements

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

## MODIFIED Requirements

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

(Previously: Default values were the responsibility of `RecommendationConfig`. These two fields already existed on the DTO pre-change — required by `PublicationVerdictEvaluator` — and were omitted from the original delta table by oversight.)

> **Rationale**: Defaults in domain DTOs introduce hidden dependencies on infrastructure decisions. Defaults belong in `EnvConfig` (infrastructure config layer) and are resolved at wiring time.

#### Scenario: RecommendationSettingsDTO is frozen and has no defaults

- GIVEN the `RecommendationSettingsDTO` class
- WHEN it is instantiated without arguments
- THEN Python raises a `TypeError` due to missing required arguments

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

(Previously: Read environment variables directly within wiring methods using `os.getenv` and loaded recommendation settings via `RecommendationConfig`.)

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

---

## REMOVED Requirements

### Requirement: RecommendationConfig Infrastructure Config Class

(Reason: Replaced by centralized `EnvConfig` configuration class.)
(Migration: Use `EnvConfig` inside `AnalyzeDocumentUseCaseWiring`.)
