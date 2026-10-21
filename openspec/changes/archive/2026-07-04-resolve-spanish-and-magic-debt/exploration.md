## Exploration: resolve-spanish-and-magic-debt

### Current State

The codebase has been audited across all directories under `src/` to identify variables, functions, class names, constants, and magic literals in Spanish or containing undocumented magic values.

#### 1. Spanish Identifiers
The production code is largely in English, following the clean architecture guidelines. However, we identified one non-exempt enum utilizing Spanish:
- **`ArticleSize`** ([article_size.py](file:///E:/Python/silvina-editorial/src/domain/enums/article_size.py)): Members and values are in Spanish (`LARGO = "largo"`, `CORTO = "corto"`, `NO_DEFINIDO = "no_definido"`, and `FUERA_RANGO = "fuera_rango"`). This enum represents internal character-count size classifications rather than verbatim document headers/LLM labels and is not listed in the deferred Spanish enums exceptions.

Additionally, several test method names use Spanish words, primarily because they test Spanish domain keywords or exceptions:
- `test_aceptable_keyword_infers_six_point_zero`
- `test_deficiente_keyword_infers_four_point_zero`
- `test_fuentes_bibliograficas_maps_to_referencias`
- `test_argumentacion_and_conclusiones_always_come_from_call_two`
- `test_argumentacion_block_is_not_misclassified_as_claridad`
- `test_missing_argumentacion_is_invalid`
- `test_claridad_and_coherencia_always_come_from_call_one`
- `test_cientifico_returns_7_sections`
- `test_desarrollo_not_in_cientifico`
- `test_desarrollo_not_in_opinion`
- `test_missing_desarrollo_is_invalid`
- `test_all_divulgacion_sections_present_is_valid`
- `test_divulgacion_returns_5_sections`
- `test_repositorio_no_violations` (testing matching of "(repositorio trazable)")
- `test_s04_repositorio_prefix_is_non_author` (testing matching of "repositorio")
- `test_english_alias_abstract_maps_to_resumen`
- `test_missing_resumen_is_invalid`

#### 2. Hardcoded Magic Values
Several production modules contain hardcoded magic numbers or strings that should be replaced with named constants:
- **`ArticleSizeClassifier`** ([article_size_classifier.py](file:///E:/Python/silvina-editorial/src/domain/classification/article_size_classifier.py)): Character limits `16000`, `24000`, `24001`, `35999`, `36000`, and `40000` are hardcoded in `classify()`.
- **`QualityLevelResolver`** ([quality_level_resolver.py](file:///E:/Python/silvina-editorial/src/domain/quality/quality_level_resolver.py)): Quality score thresholds `9.0`, `7.0`, `5.0`, and `3.0` are hardcoded in `resolve()`.
- **`PublicationVerdictEvaluator`** ([publication_verdict_evaluator.py](file:///E:/Python/silvina-editorial/src/domain/recommendation/publication_verdict_evaluator.py)): Thresholds `5.0` (quality score) and `5.0` (grammar score) are hardcoded in `evaluate()`.
- **`StructureValidator`** ([structure_validator.py](file:///E:/Python/silvina-editorial/src/domain/structure/structure_validator.py)): The maximum character limit for a header line (`100`) is hardcoded in `_extract_present_sections()`.
- **`DocxCitationAdapter`** ([docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_citation_adapter.py)): The maximum length of an author name (`100`) is hardcoded in `_collect_multi_author()`.
- **`DocxReferenceAdapter`** ([docx_reference_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_reference_adapter.py)): The regular expression compilation flags `2 | 16` are hardcoded (representing `re.IGNORECASE | re.DOTALL`).
- **`LanguageToolAdapter`** ([language_tool_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/grammar/language_tool_adapter.py)): The maximum number of replacements returned (`3`) is hardcoded in `_map_to_dto()`.
- **`DocxReportAdapter`** ([docx_report_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/report/docx_report_adapter.py)):
  - Divisor for page count calculation `250` (representing words per page).
  - List slice size `5` for limiting the number of displayed grammar errors and APA violations.
  - Word context truncation size `150` for displaying grammar errors.
  - Replacements limit `3` for grammar suggestions.
  - Table row size `6` and column size `2` for metrics table creation.
  - Repeat separator count `80` for the footer line separator.

### Affected Areas

The proposed changes will touch:
- **Enums**:
  - [article_size.py](file:///E:/Python/silvina-editorial/src/domain/enums/article_size.py)
- **Domain Services**:
  - [article_size_classifier.py](file:///E:/Python/silvina-editorial/src/domain/classification/article_size_classifier.py)
  - [quality_level_resolver.py](file:///E:/Python/silvina-editorial/src/domain/quality/quality_level_resolver.py)
  - [publication_verdict_evaluator.py](file:///E:/Python/silvina-editorial/src/domain/recommendation/publication_verdict_evaluator.py)
  - [structure_validator.py](file:///E:/Python/silvina-editorial/src/domain/structure/structure_validator.py)
- **Infrastructure Adapters**:
  - [docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_citation_adapter.py)
  - [docx_reference_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_reference_adapter.py)
  - [language_tool_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/grammar/language_tool_adapter.py)
  - [docx_report_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/report/docx_report_adapter.py)
- **Tests**:
  - [test_article_size.py](file:///E:/Python/silvina-editorial/src/domain/tests/enums/test_article_size.py)
  - [test_article_size_classifier.py](file:///E:/Python/silvina-editorial/src/domain/tests/classification/test_article_size_classifier.py)
  - [test_quality_level_resolver.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/test_quality_level_resolver.py)
  - [test_classification_result.py](file:///E:/Python/silvina-editorial/src/domain/tests/dtos/test_classification_result.py)

### Approaches

#### Approach 1: Complete Refactoring in a Single Pass
Refactor the Spanish `ArticleSize` enum to English (`LARGE`, `SHORT`, `UNDEFINED`, `OUT_OF_RANGE`) along with all of its references, and replace all hardcoded magic numbers/strings in production code with named class or module constants.
*   **Pros**: Completely resolves Item 4 in a clean manner.
*   **Cons**: Higher initial testing effort to verify that renaming the enum and its values doesn't break dependent serialization or display components.

#### Approach 2: Phased Refactoring
Resolve only the magic values (numbers and regex flags) first, while postponing `ArticleSize` renaming until the final pass of Spanish enums (which includes `ArticleType`, `SectionType`, etc.).
*   **Pros**: Reduces immediate diff footprint and keeps enums aligned.
*   **Cons**: Leaves `ArticleSize` (an internal computed classification enum) in Spanish, which violates Clean Architecture guidelines for non-exempt structures.

### Recommendation

We recommend **Approach 1**. `ArticleSize` is an internal classification calculated from character counts; it is not dependent on matching literal LLM output or document headers (unlike `ArticleType` and `SectionType`), and thus it should be standard English. Promoting all magic literals to named constants in the same pass ensures the technical debt for Item 4 is fully paid.

### Risks

- **Enum renaming**: Renaming the values `"largo"`, `"corto"`, `"no_definido"`, `"fuera_rango"` to `"large"`, `"short"`, `"undefined"`, `"out_of_range"` could affect formatting/display templates or serialization logic if any exist. This will be mitigated by doing a global project-wide search and updating the Word report template rendering code.
- **Regex flags**: Changing raw flags `2 | 16` to `re.IGNORECASE | re.DOTALL` requires importing the standard flags. This is low risk but needs confirmation through test suites.

### Ready for Proposal
Yes
