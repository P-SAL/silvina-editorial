# Design: Resolve Slice 5 Technical Debt

This change addresses deferred OOP and encapsulation technical debt items from Slice 5 in `openspec/TECHNICAL_DEBT.md` to resolve signature drift and align with hexagonal architecture standards.

## Module Layout
```text
src/
├── domain/
│   ├── enums/
│   │   └── quality_level.py                            # MODIFIED - Removed get_quality_level_from_score()
│   ├── quality/
│   │   ├── quality_level_resolver.py                   # NEW - Contains QualityLevelResolver domain service
│   │   ├── quality_analyzer.py                         # MODIFIED - Added optional resolver dependency
│   │   └── quality_text_sampler.py                     # MODIFIED - Encapsulated _CONCLUSION_HEADER_PATTERN
│   └── tests/
│       ├── enums/
│       │   └── test_get_quality_level_from_score.py    # DELETED - Ported to test_quality_level_resolver.py
│       └── quality/
│           ├── test_quality_level_resolver.py          # NEW - Unit tests for QualityLevelResolver
│           └── test_quality_text_sampler.py            # MODIFIED - Encapsulated build_document_content()
├── application/
│   └── tests/
│       ├── fake_llm_generator_adapter.py               # MODIFIED - Added options parameter to generate()
│       └── test_analyze_quality_use_case.py            # MODIFIED - Updated setUp QualityAnalyzer instantiation
└── infrastructure/
    └── wirings/
        └── analyze_quality_use_case_wiring.py          # MODIFIED - Injected QualityLevelResolver into QualityAnalyzer
```

## Architecture Decisions

### 1. Extraction of QualityLevelResolver
The helper function `get_quality_level_from_score()` in `src/domain/enums/quality_level.py` is refactored into a domain service class `QualityLevelResolver` in `src/domain/quality/quality_level_resolver.py`.
```python
from src.domain.enums.quality_level import QualityLevel

class QualityLevelResolver:
    """Domain service mapping overall score to QualityLevel."""

    def resolve(self, score: float) -> QualityLevel:
        if score >= 9.0:
            return QualityLevel.EXCELLENT
        if score >= 7.0:
            return QualityLevel.GOOD
        if score >= 5.0:
            return QualityLevel.ACCEPTABLE
        if score >= 3.0:
            return QualityLevel.NEEDS_IMPROVEMENT
        return QualityLevel.POOR
```

### 2. Optional Constructor Injection in QualityAnalyzer
To support clean dependency injection while preventing breakages at legacy instantiation sites (e.g., in other unit tests), the `QualityAnalyzer` constructor accepts `resolver: QualityLevelResolver | None = None` defaulting to `None`. If `None`, it falls back to creating a `QualityLevelResolver()` instance.
```python
class QualityAnalyzer:
    def __init__(
        self,
        llm_generator: LlmGeneratorPort,
        text_sampler: QualityTextSampler,
        response_parser: QualityResponseParser,
        clarity_coherence_prompt_template: str,
        argumentation_conclusions_prompt_template: str,
        resolver: QualityLevelResolver | None = None,
    ) -> None:
        ...
        self._resolver = resolver or QualityLevelResolver()
```

### 3. Encapsulating regex pattern in QualityTextSampler
The module-level variable `_CONCLUSION_HEADER_PATTERN` is refactored into a private class-level attribute inside `QualityTextSampler`.
```python
class QualityTextSampler:
    _CONCLUSION_HEADER_PATTERN = re.compile(r"conclusi", re.IGNORECASE)
    ...
```

### 4. Encapsulating build_document_content helper in test case
The standalone function `build_document_content()` in `src/domain/tests/quality/test_quality_text_sampler.py` is moved inside `TestQualityTextSampler` as a private helper method `_build_document_content()`.

### 5. Signature Alignment in FakeLlmGeneratorAdapterForTest
The signature of `FakeLlmGeneratorAdapterForTest.generate` in `src/application/tests/fake_llm_generator_adapter.py` is updated to include the optional parameter `options: dict | None = None` to match the `LlmGeneratorPort` contract.

## Test Strategy
1. Delete `src/domain/tests/enums/test_get_quality_level_from_score.py`.
2. Create `src/domain/tests/quality/test_quality_level_resolver.py` testing `QualityLevelResolver.resolve()` using the exact boundary assertions from the deleted test.
3. Run the complete `pytest` test suite to ensure zero regressions.
