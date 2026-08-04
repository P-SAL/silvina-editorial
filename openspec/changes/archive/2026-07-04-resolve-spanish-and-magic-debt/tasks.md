# Tasks: Resolve Spanish and Magic Debt

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: stacked-to-main
400-line budget risk: Low

## Phase 1: Foundation
- [x] **Define Env Vars**: Add variables to [.env](file:///E:/Python/silvina-editorial/.env) and [.env.example](file:///E:/Python/silvina-editorial/.env.example):
  - `ARTICLE_SIZE_SHORT_MIN_CHARS=16000`, `ARTICLE_SIZE_SHORT_MAX_CHARS=24000`
  - `ARTICLE_SIZE_UNDEFINED_MIN_CHARS=24001`, `ARTICLE_SIZE_UNDEFINED_MAX_CHARS=35999`
  - `ARTICLE_SIZE_LONG_MIN_CHARS=36000`, `ARTICLE_SIZE_LONG_MAX_CHARS=40000`
  - `QUALITY_LEVEL_EXCELLENT_THRESHOLD=9.0`, `QUALITY_LEVEL_GOOD_THRESHOLD=7.0`, `QUALITY_LEVEL_ACCEPTABLE_THRESHOLD=5.0`, `QUALITY_LEVEL_NEEDS_IMPROVEMENT_THRESHOLD=3.0`
  - `RECOMMENDATION_CRITICAL_QUALITY_THRESHOLD=5.0`, `RECOMMENDATION_CRITICAL_GRAMMAR_THRESHOLD=5.0`
  - Added manually by the user (grouped by category, with English comments) since `.env*` paths are hard-denied by session permission settings.
- [x] **Update Settings DTO**: Add `critical_quality_threshold` and `critical_grammar_threshold` fields to `RecommendationSettingsDTO` in [recommendation_settings_dto.py](file:///E:/Python/silvina-editorial/src/domain/dtos/recommendation_settings_dto.py).
- [x] **Update RecommendationConfig**: In [recommendation_config.py](file:///E:/Python/silvina-editorial/src/infrastructure/config/recommendation_config.py), load the new critical thresholds and map them in `build_settings`. Clean up `os` imports to `from os import getenv` standard.

## Phase 2: Core Refactoring
- [x] **Refactor ArticleSize Enum**: In [article_size.py](file:///E:/Python/silvina-editorial/src/domain/enums/article_size.py), rename members to `LONG`, `SHORT`, `UNDEFINED`, `OUT_OF_RANGE`. Keep the Spanish string values.
- [x] **Update ArticleClassifier**: In [article_classifier.py](file:///E:/Python/silvina-editorial/src/domain/classification/article_classifier.py), replace `ArticleSize.FUERA_RANGO` with `ArticleSize.OUT_OF_RANGE`.
- [x] **Parameterize ArticleSizeClassifier**: In [article_size_classifier.py](file:///E:/Python/silvina-editorial/src/domain/classification/article_size_classifier.py), accept keyword-only thresholds in constructor with current default values, and use them in `classify()`. Return the new enum members.
- [x] **Parameterize QualityLevelResolver**: In [quality_level_resolver.py](file:///E:/Python/silvina-editorial/src/domain/quality/quality_level_resolver.py), accept keyword-only tier thresholds in constructor (default: 9.0, 7.0, 5.0, 3.0), and use them in `resolve()`.
- [x] **Parameterize PublicationVerdictEvaluator**: In [publication_verdict_evaluator.py](file:///E:/Python/silvina-editorial/src/domain/recommendation/publication_verdict_evaluator.py), replace hardcoded `5.0` limits with `context.settings.critical_quality_threshold` and `context.settings.critical_grammar_threshold`.

## Phase 3: Infrastructure Wiring
- [x] **Update ClassifyArticleUseCaseWiring**: In [classify_article_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/classify_article_use_case_wiring.py), load and pass character count limits to `ArticleSizeClassifier`.
- [x] **Update AnalyzeQualityUseCaseWiring**: In [analyze_quality_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_quality_use_case_wiring.py), load and pass tier thresholds to `QualityLevelResolver`.

## Phase 4: Unit Test Alignment
- [x] **Update Enum and Classifier Tests**: Modify assertions in [test_article_size.py](file:///E:/Python/silvina-editorial/src/domain/tests/enums/test_article_size.py) and [test_article_size_classifier.py](file:///E:/Python/silvina-editorial/src/domain/tests/classification/test_article_size_classifier.py) to use new English enum members and test custom boundaries.
- [x] **Update Quality Resolver Tests**: Add test cases for custom thresholds in [test_quality_level_resolver.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/test_quality_level_resolver.py).
- [x] **Update DTO and Recommendation Tests**: Update [test_analysis_result.py](file:///E:/Python/silvina-editorial/src/domain/tests/dtos/test_analysis_result.py), [test_classification_result.py](file:///E:/Python/silvina-editorial/src/domain/tests/dtos/test_classification_result.py), and [test_recommendation_builder.py](file:///E:/Python/silvina-editorial/src/domain/tests/recommendation/test_recommendation_builder.py) to align with modified settings, enum members, and custom settings assertion.
- [x] **Update Orchestrator Wiring Tests**: Adjust [test_analyze_document_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_analyze_document_use_case_wiring.py) default settings test to assert the critical thresholds.
- [x] **Update Legacy Root Test Suites** (discovered during apply, not in original task list): `tests/test_main_dto_mapping.py`, `tests/test_main_cli_args.py`, `tests/e2e/test_gradio_e2e.py` also construct `ArticleSize.CORTO` fixtures — updated to `ArticleSize.SHORT`.

## Phase 5: Verification
- [x] **Run Unit Tests**: Execute `.venv\Scripts\pytest` to verify all 635+ tests pass successfully. Result: 641 passed, 3 skipped, 6 subtests passed (635 baseline + 6 new tests added for custom-threshold coverage).
- [x] **Manual CLI Verification**: Ran a quick classification/quality sanity check via the wirings directly (no live Ollama server needed) — confirmed `ArticleSizeClassifier`/`QualityLevelResolver` produce correct enum results with defaults, and `ArticleSize` members still serialize to the original Spanish string values (`corto`, `largo`, `fuera_rango`).
