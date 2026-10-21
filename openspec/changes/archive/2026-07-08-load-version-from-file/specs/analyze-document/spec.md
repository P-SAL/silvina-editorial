# Delta for analyze-document

## MODIFIED Requirements

### Requirement: EnvConfig Infrastructure Config Class

`EnvConfig` MUST reside in `src/infrastructure/env_config.py`. It MUST parse environment variables at instantiation, cast them, and cache them as typed instance attributes. It MUST expose a method `get_recommendation_settings() -> RecommendationSettingsDTO` to build recommendation settings.

(Previously: SILVINA_VERSION was loaded from environment variables with a default of "0.9".)

The application version attribute (`silvina_version`) MUST be resolved dynamically:
- In production/standard mode: `EnvConfig` MUST load the version string from the file `version.txt` located in the project root directory (resolved relative to `EnvConfig` file location: `Path(__file__).resolve().parents[2] / "version.txt"`). The version string MUST be stripped of surrounding whitespace. If the file is missing or unreadable, `EnvConfig` MUST raise `FileNotFoundError` (or standard OS/permission errors).
- In testing mode (when the environment variable `TESTING` is `"True"`, `"true"`, or `"1"`): `EnvConfig` MUST fall back to loading the version from the environment variable `SILVINA_VERSION` (defaulting to `"0.9"` if the variable is not set), without requiring the `version.txt` file to exist.

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
| `REPORT_SCORE_HIGH_THRESHOLD` | `float` | `8.0` | `report_score_high_threshold` |
| `REPORT_SCORE_MEDIUM_THRESHOLD` | `float` | `6.0` | `report_score_medium_threshold` |
| `REPORT_WORDS_PER_PAGE` | `int` | `250` | `report_words_per_page` |
| `REPORT_MAX_ERRORS_DISPLAYED` | `int` | `5` | `report_max_errors_displayed` |
| `REPORT_CONTEXT_TRUNCATION_LIMIT` | `int` | `150` | `report_context_truncation_limit` |
| `REPORT_MAX_REPLACEMENTS` | `int` | `3` | `report_max_replacements` |

> **Naming note**: the `PUBLISH_THRESHOLD` … `CRITICAL_GRAMMAR_THRESHOLD` variables (recommendation thresholds) carry no `RECOMMENDATION_` prefix. In `.env`/`.env.example` they MUST be grouped under a section comment (e.g. `# Recommendation thresholds`) instead of relying on a name prefix for grouping.

#### Scenario: EnvConfig defaults are loaded when env is empty and version.txt exists

- GIVEN an empty environment except for a valid `version.txt` file with content `"0.95"`
- WHEN `EnvConfig` is instantiated
- THEN attributes match the defaults in the table above
- AND `env_config.silvina_version` is `"0.95"`

#### Scenario: EnvConfig parses and casts environment variables

- GIVEN environment contains `CITATION_MAX_AUTHOR_NAME_LENGTH=150`
- AND a valid `version.txt` file exists
- WHEN `EnvConfig` is instantiated
- THEN `env_config.citation_max_author_name_length == 150`

#### Scenario: EnvConfig fails fast when version.txt is missing

- GIVEN the `version.txt` file is missing in the root directory
- AND `TESTING` environment variable is not set to `"True"`, `"true"`, or `"1"`
- WHEN `EnvConfig` is instantiated
- THEN a `FileNotFoundError` is raised

#### Scenario: EnvConfig falls back to environment variable in testing mode

- GIVEN `TESTING` environment variable is set to `"True"`
- AND `SILVINA_VERSION` environment variable is set to `"0.99"`
- AND `version.txt` is missing
- WHEN `EnvConfig` is instantiated
- THEN `env_config.silvina_version` is `"0.99"`

#### Scenario: EnvConfig uses testing default version when not specified in env

- GIVEN `TESTING` environment variable is set to `"True"`
- AND `SILVINA_VERSION` environment variable is not set
- AND `version.txt` is missing
- WHEN `EnvConfig` is instantiated
- THEN `env_config.silvina_version` is `"0.9"`
