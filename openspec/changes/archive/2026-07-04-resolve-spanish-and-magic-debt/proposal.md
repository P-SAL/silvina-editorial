# Proposal: Resolve Spanish and Magic Debt

## Intent
Resolve Technical Debt Item 4 by refactoring variables, functions, and constants in Spanish to English, and replacing hardcoded magic numbers with named constants configured via environment variables.

## Scope
- Refactor members of `ArticleSize` enum in [article_size.py](file:///E:/Python/silvina-editorial/src/domain/enums/article_size.py) to English (`LONG`, `SHORT`, `UNDEFINED`, `OUT_OF_RANGE`) while preserving their Spanish string values (`"largo"`, `"corto"`, `"no_definido"`, `"fuera_rango"`) to keep report output and downstream mappings intact.
- Audit and replace magic numbers in:
  - [article_size_classifier.py](file:///E:/Python/silvina-editorial/src/domain/classification/article_size_classifier.py) (character count thresholds).
  - [quality_level_resolver.py](file:///E:/Python/silvina-editorial/src/domain/quality/quality_level_resolver.py) (quality tier thresholds).
  - [publication_verdict_evaluator.py](file:///E:/Python/silvina-editorial/src/domain/recommendation/publication_verdict_evaluator.py) (critical quality and grammar limits).
- Load these thresholds from `.env` inside their respective wirings/configs and inject them.

## Approach
1. **Refactor `ArticleSize` Enum**: Rename members to `LONG`, `SHORT`, `UNDEFINED`, and `OUT_OF_RANGE`. Keep values as `"largo"`, `"corto"`, `"no_definido"`, and `"fuera_rango"`.
2. **Inject Classifier Thresholds**: Modify `ArticleSizeClassifier`'s constructor to accept character count ranges (min/max for short, undefined, long). Update `ClassifyArticleUseCaseWiring` to load these values from environment variables (e.g., `ARTICLE_SIZE_SHORT_MIN_CHARS`, etc.) and inject them.
3. **Inject Quality Resolver Thresholds**: Modify `QualityLevelResolver`'s constructor to accept thresholds for each quality level. Update `AnalyzeQualityUseCaseWiring` to load them from environment variables (e.g., `QUALITY_LEVEL_EXCELLENT_THRESHOLD`, etc.) and inject them.
4. **Inject Verdict Evaluator Thresholds**: Add `critical_quality_threshold` and `critical_grammar_threshold` to `RecommendationSettingsDTO` and `RecommendationConfig` (loaded from `.env`). Update `PublicationVerdictEvaluator` to evaluate using `context.settings` instead of the hardcoded `5.0`.
5. **Update All References**: Update all references to `ArticleSize` members in domain files, DTOs, and test suites.
6. **Verification**: Run `pytest` to ensure all 635 tests pass with no regressions.

## Capabilities
No new capabilities are introduced. This is a refactoring change to resolve technical debt.
