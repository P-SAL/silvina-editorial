# Technical Design: Resolve Spanish and Magic Debt

## Technical Approach
We will refactor the `ArticleSize` enum members from Spanish to English while keeping their underlying Spanish string values intact to prevent breaking downstream reporting or CLI display behaviors. Additionally, we will remove hardcoded magic numbers from the domain classification, quality, and verdict evaluation layers, replacing them with constructor-injected thresholds loaded from environment variables in their respective wirings and configurations.

### 1. Refactor `ArticleSize` Enum members to English
Refactor `src/domain/enums/article_size.py` members:
- `LARGO` -> `LONG`
- `CORTO` -> `SHORT`
- `NO_DEFINIDO` -> `UNDEFINED`
- `FUERA_RANGO` -> `OUT_OF_RANGE`

Keep values as `"largo"`, `"corto"`, `"no_definido"`, and `"fuera_rango"`.

### 2. Parameterize Classifiers and Resolvers
- **`ArticleSizeClassifier`**: Inject ranges (six boundaries: minimum/maximum for short, long, undefined) with default values matching the current boundaries to preserve backward compatibility.
- **`QualityLevelResolver`**: Inject tier thresholds (four thresholds: excellent, good, acceptable, needs improvement) with default values mapping to `9.0`, `7.0`, `5.0`, `3.0` respectively.
- **`PublicationVerdictEvaluator`**: Update to use `context.settings.critical_quality_threshold` and `context.settings.critical_grammar_threshold` instead of hardcoded `5.0`.

### 3. Wiring and Config Updates
- **`ClassifyArticleUseCaseWiring`**: Load classifier thresholds via `getenv` and pass them to the `ArticleSizeClassifier` constructor.
- **`AnalyzeQualityUseCaseWiring`**: Load quality level thresholds via `getenv` and pass them to the `QualityLevelResolver` constructor.
- **`RecommendationConfig`**: Load new threshold configurations `critical_quality_threshold` and `critical_grammar_threshold` from environment variables and bind them to `RecommendationSettingsDTO`. Align `os` imports to the specific name import standard `from os import getenv` to comply with Clean Architecture rules.

---

## Architecture Decisions

1. **Member-Only Translation**: Keeping Spanish string values preserves CLI, serialization, and report display behavior, eliminating any downstream UI or serialization risks.
2. **Constructor Injection**: Passing configuration values at construction time keeps domain services pure and decoupled from infrastructure-specific environment lookups (`os.getenv`), adhering strictly to hexagonal boundaries.
3. **Keyword-Only Defaults**: Using keyword-only constructor parameters with default values prevents breaking other tests or instantiation points that do not yet use the parameterized configuration.

---

## File Changes

### Domain Layer
- **[article_size.py](file:///E:/Python/silvina-editorial/src/domain/enums/article_size.py)**: Rename enum members to `LONG`, `SHORT`, `UNDEFINED`, and `OUT_OF_RANGE`.
- **[article_size_classifier.py](file:///E:/Python/silvina-editorial/src/domain/classification/article_size_classifier.py)**: Implement `__init__` with parameterized character limits, update `classify` logic and return enum members.
- **[quality_level_resolver.py](file:///E:/Python/silvina-editorial/src/domain/quality/quality_level_resolver.py)**: Implement `__init__` with parameterized tier thresholds, update `resolve` logic.
- **[recommendation_settings_dto.py](file:///E:/Python/silvina-editorial/src/domain/dtos/recommendation_settings_dto.py)**: Add `critical_quality_threshold` and `critical_grammar_threshold`.
- **[publication_verdict_evaluator.py](file:///E:/Python/silvina-editorial/src/domain/recommendation/publication_verdict_evaluator.py)**: Replace hardcoded quality/grammar checks (`5.0`) with `context.settings` checks.
- **[article_classifier.py](file:///E:/Python/silvina-editorial/src/domain/classification/article_classifier.py)**: Update references from `ArticleSize.FUERA_RANGO` to `ArticleSize.OUT_OF_RANGE`.

### Infrastructure Layer
- **[recommendation_config.py](file:///E:/Python/silvina-editorial/src/infrastructure/config/recommendation_config.py)**: Read critical quality and grammar thresholds from `.env` and map them to `RecommendationSettingsDTO`. Correct import style.
- **[classify_article_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/classify_article_use_case_wiring.py)**: Load and inject `ArticleSizeClassifier` character count bounds.
- **[analyze_quality_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_quality_use_case_wiring.py)**: Load and inject `QualityLevelResolver` quality thresholds.
- **[.env](file:///E:/Python/silvina-editorial/.env)** & **[.env.example](file:///E:/Python/silvina-editorial/.env.example)**: Declare new variables with default values.

### Test Suites
- **[test_article_size.py](file:///E:/Python/silvina-editorial/src/domain/tests/enums/test_article_size.py)**: Update assertions to test English enum members.
- **[test_article_size_classifier.py](file:///E:/Python/silvina-editorial/src/domain/tests/classification/test_article_size_classifier.py)**: Test parameterized initialization and update enum members.
- **[test_quality_level_resolver.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/test_quality_level_resolver.py)**: Test parameterized initialization.
- **[test_recommendation_builder.py](file:///E:/Python/silvina-editorial/src/domain/tests/recommendation/test_recommendation_builder.py)**: Add settings DTO values, verify dynamic threshold evaluation.
- **[test_analysis_result.py](file:///E:/Python/silvina-editorial/src/domain/tests/dtos/test_analysis_result.py)** & **[test_classification_result.py](file:///E:/Python/silvina-editorial/src/domain/tests/dtos/test_classification_result.py)**: Update `ArticleSize` enum member usages.

---

## Testing Strategy
1. **Unit Tests**:
   - Verify `ArticleSizeClassifier` constructor properly overrides defaults when given specific values, and maps correctly on bounds.
   - Verify `QualityLevelResolver` constructor correctly maps custom scores to quality levels.
   - Verify `PublicationVerdictEvaluator` applies `critical_quality_threshold` and `critical_grammar_threshold` dynamically from the DTO.
2. **Integration Tests**:
   - Verify wirings instantiate objects with values loaded from the environment.
3. **Regression Tests**:
   - Execute complete suite (`pytest`) to verify all 630+ tests continue to pass.
