# Tasks: Resolve Slice 5 Technical Debt

## Review Workload Forecast
- **Expected Change Size**: ~120 lines total
- **Risk Level**: Low
- **Strategy**: Single PR (`single-pr`)
- **Metadata**:
Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Low

---

## Phase 1: Foundation

### T-1.1: Create QualityLevelResolver [x]
- Create `src/domain/quality/quality_level_resolver.py` containing `QualityLevelResolver` with a `resolve(self, score: float) -> QualityLevel` method.
- Remove `get_quality_level_from_score` function from `src/domain/enums/quality_level.py`.

### T-1.2: Port and Migrate Quality Level Tests [x]
- Create `src/domain/tests/quality/test_quality_level_resolver.py`.
- Port all test cases from `src/domain/tests/enums/test_get_quality_level_from_score.py` to target `QualityLevelResolver.resolve()`.
- Delete `src/domain/tests/enums/test_get_quality_level_from_score.py`.

### T-1.3: Update Fake LLM Generator Method Signature [x]
- Modify `src/application/tests/fake_llm_generator_adapter.py` to add `options: dict | None = None` to `FakeLlmGeneratorAdapterForTest.generate`.

---

## Phase 2: Core Refactoring

### T-2.1: Inject QualityLevelResolver in QualityAnalyzer [x]
- Import `QualityLevelResolver` in `src/domain/quality/quality_analyzer.py`.
- Add optional `resolver: QualityLevelResolver | None = None` parameter to constructor, fallback to default instance.
- Replace direct `get_quality_level_from_score(overall_score)` call with `self._resolver.resolve(overall_score)`.

### T-2.2: Encapsulate Regex in QualityTextSampler [x]
- Define `_CONCLUSION_HEADER_PATTERN` inside the `QualityTextSampler` class in `src/domain/quality/quality_text_sampler.py`.
- Update internal references from `_CONCLUSION_HEADER_PATTERN` to `self._CONCLUSION_HEADER_PATTERN`.

### T-2.3: Encapsulate Document Content Helper in Test Case [x]
- Move `build_document_content` in `src/domain/tests/quality/test_quality_text_sampler.py` inside the `TestQualityTextSampler` class as `_build_document_content`.
- Update all calls to use `self._build_document_content`.

---

## Phase 3: Wiring

### T-3.1: Update Use Case Wiring [x]
- In `src/infrastructure/wirings/analyze_quality_use_case_wiring.py`, import `QualityLevelResolver` and instantiate it, passing it to the `QualityAnalyzer` constructor.

### T-3.2: Update Application Tests Setup [x]
- In `src/application/tests/test_analyze_quality_use_case.py`, update `QualityAnalyzer` setup to construct it with a default resolver if necessary.

---

## Phase 4: Verification

### T-4.1: Execute Tests [x]
- Run `pytest` to verify all tests pass, including the new `test_quality_level_resolver.py`.
