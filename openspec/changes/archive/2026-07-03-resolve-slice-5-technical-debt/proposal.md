# Proposal: Resolve Slice 5 Technical Debt

## Intent
Address four deferred OOP and encapsulation technical debt items from Slice 5 (Item 3 in `openspec/TECHNICAL_DEBT.md`) to resolve signature drift and improve clean architecture adherence.

## Scope
### In Scope
- Refactoring loose domain helper `get_quality_level_from_score()` into `QualityLevelResolver.resolve()`.
- Injecting `QualityLevelResolver` as an optional constructor dependency in `QualityAnalyzer`.
- Encapsulating regex pattern `_CONCLUSION_HEADER_PATTERN` within the `QualityTextSampler` class.
- Encapsulating test helper `build_document_content()` as a private method in `TestQualityTextSampler`.
- Updating `FakeLlmGeneratorAdapterForTest.generate()` method signature to accept `options: dict | None = None`.

### Out of Scope
- Modifying any other LLM generator fakes (e.g., `FakeLlmGeneratorAdapter`).
- Changing the mapping logic inside `QualityLevelResolver`.
- Refactoring other items in the Technical Debt Registry.

## Capabilities
- **New Capabilities**: None
- **Modified Capabilities**: None

## Approach
- **Optional Constructor Injection**: Inject `resolver: QualityLevelResolver | None = None` into `QualityAnalyzer`. Default to `QualityLevelResolver()` if `None` to prevent breaking existing instantiation sites (wiring and tests).
- **Encapsulate Resolution Logic**: Implement `QualityLevelResolver` in a new domain service with exact mapping logic of `get_quality_level_from_score()`.
- **Encapsulate Patterns & Helpers**: Move `_CONCLUSION_HEADER_PATTERN` inside `QualityTextSampler` as a private class attribute. Move `build_document_content()` inside `TestQualityTextSampler` as `_build_document_content()`.
- **Signature Alignment**: Add `options: dict | None = None` to `FakeLlmGeneratorAdapterForTest.generate()`.

## Affected Areas
| Path (relative to repo root) | Impact |
| :--- | :--- |
| `src/domain/enums/quality_level.py` | Remove helper function `get_quality_level_from_score()`. |
| `src/domain/quality/quality_level_resolver.py` | Create domain service `QualityLevelResolver`. |
| `src/domain/quality/quality_analyzer.py` | Inject and use `QualityLevelResolver`. |
| `src/infrastructure/wirings/analyze_quality_use_case_wiring.py` | Wire `QualityLevelResolver` to `QualityAnalyzer`. |
| `src/domain/quality/quality_text_sampler.py` | Encapsulate regex pattern inside `QualityTextSampler`. |
| `src/domain/tests/quality/test_quality_text_sampler.py` | Encapsulate helper inside `TestQualityTextSampler` class. |
| `src/application/tests/fake_llm_generator_adapter.py` | Update `FakeLlmGeneratorAdapterForTest.generate()` signature. |
| `src/domain/tests/quality/test_quality_level_resolver.py` | Create to test `QualityLevelResolver`. (Ported from `test_get_quality_level_from_score.py`). |
| `src/domain/tests/enums/test_get_quality_level_from_score.py` | Delete file. |

## Risks & Mitigation
- **Risk**: Missing dependency wiring breaks instantiation of `QualityAnalyzer`.
- **Mitigation**: Optional parameter defaulting to `QualityLevelResolver()` and full test suite verification.

## Rollback Plan
- Revert changes via Git checkout of the affected files.

## Success Criteria
- Test suite passes cleanly with zero errors.
- Verification of new `QualityLevelResolver` functionality.
