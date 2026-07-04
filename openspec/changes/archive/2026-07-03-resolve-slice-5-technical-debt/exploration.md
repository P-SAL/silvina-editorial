## Exploration: Slice 5 Technical Debt Cleanup

### Current State
Today, several technical debt items from Slice 5 (Item 3 in `openspec/TECHNICAL_DEBT.md`) violate encapsulation, object-oriented programming guidelines, or interface signatures established in later slices:
1. `get_quality_level_from_score()` is a module-level helper function in `src/domain/enums/quality_level.py`. Clean Hexagonal Architecture conventions require domain helpers to be domain service classes (such as `ArticleSizeClassifier`) rather than loose functions.
2. `_CONCLUSION_HEADER_PATTERN` is defined at the module scope in `src/domain/quality/quality_text_sampler.py`, leaking regex patterns outside the class context.
3. `build_document_content()` in `src/domain/tests/quality/test_quality_text_sampler.py` is a module-level helper function instead of a private test case helper method.
4. `FakeLlmGeneratorAdapterForTest.generate()` in `src/application/tests/fake_llm_generator_adapter.py` lacks the `options: dict | None = None` parameter defined on `LlmGeneratorPort`, creating a signature drift.

### Affected Areas
- `src/domain/enums/quality_level.py` — The helper function `get_quality_level_from_score()` should be removed.
- `src/domain/quality/quality_level_resolver.py` — Needs to be created as a new domain service containing the class `QualityLevelResolver` and its public method `resolve(self, score: float) -> QualityLevel`.
- `src/domain/quality/quality_analyzer.py` — Needs to import and inject `QualityLevelResolver` in its constructor, updating the mapping call.
- `src/infrastructure/wirings/analyze_quality_use_case_wiring.py` — Needs to instantiate and wire `QualityLevelResolver` into the `QualityAnalyzer`.
- `src/domain/tests/quality/test_quality_analyzer.py` — Needs to instantiate `QualityLevelResolver` in helpers that construct `QualityAnalyzer`.
- `src/application/tests/test_analyze_quality_use_case.py` — Needs to construct `QualityAnalyzer` with the new dependency.
- `src/domain/tests/enums/test_get_quality_level_from_score.py` — Needs to be moved to `src/domain/tests/quality/test_quality_level_resolver.py` and updated to test `QualityLevelResolver.resolve()`.
- `src/domain/quality/quality_text_sampler.py` — The regex `_CONCLUSION_HEADER_PATTERN` must be moved inside the `QualityTextSampler` class as a private class attribute.
- `src/domain/tests/quality/test_quality_text_sampler.py` — The `build_document_content()` helper must be refactored into a private method `_build_document_content()` inside the `TestQualityTextSampler` class.
- `src/application/tests/fake_llm_generator_adapter.py` — The `generate()` method signature must be updated to accept `options: dict | None = None`.

### Approaches
1. **Standard OOP Encapsulation and Dependency Injection** — Extract the resolver to `QualityLevelResolver` inside `src/domain/quality/quality_level_resolver.py`, inject it into `QualityAnalyzer`, encapsulate the regex inside `QualityTextSampler` as a private class attribute, encapsulate the test helper inside `TestQualityTextSampler` as a private method, and update `FakeLlmGeneratorAdapterForTest.generate()` signature to match the port interface.
   - Pros:
     - Fully adheres to clean architecture guidelines (classes over module-level functions).
     - Minimizes global variables and improves encapsulation.
     - Resolves the signature drift between the test double and the `LlmGeneratorPort` interface, preventing future test breakages.
     - Consistent with other classifiers in the codebase (e.g. `ArticleSizeClassifier`).
   - Cons:
     - Requires modifying several files that instantiate `QualityAnalyzer` (wiring and tests) due to the new constructor parameter.
   - Effort: Low

2. **Class-based Helpers without Constructor Injection** — Implement `QualityLevelResolver` with a class method or instantiate it directly inside `QualityAnalyzer` without constructor injection. Keep other encapsulation and signature alignment changes as-is.
   - Pros:
     - Avoids adding a parameter to `QualityAnalyzer.__init__`, reducing changes to wiring and tests.
   - Cons:
     - Violates dependency injection best practices and makes unit testing `QualityAnalyzer` harder if the resolver needs mock behavior.
     - Inconsistent with how `ArticleSizeClassifier` is handled.
   - Effort: Low

### Recommendation
We recommend **Approach 1: Standard OOP Encapsulation and Dependency Injection**. It is clean, respects the project's hexagonal guidelines, ensures high testability, and mirrors how other domain classification logic is implemented in the codebase.

### Risks
- **Test Suite Updates**: Any missing constructor update for `QualityAnalyzer` in the tests or wirings will lead to runtime errors (e.g., `TypeError: __init__() missing 1 required positional argument`). This is mitigated by running `pytest` immediately after application.
- **Import Changes**: Moving files and deleting module-level functions might lead to broken imports in legacy files if they were referencing `get_quality_level_from_score()`. However, the legacy files `business_logic/quality_analyzer.py` and `domain/enums.py` are already self-contained or import from the legacy `domain` package, so they will not be broken.

### Ready for Proposal
Yes
