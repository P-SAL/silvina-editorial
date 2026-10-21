<Tasks: Remove Unused Article Type Parameter>
## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~20 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | single-pr |
| Chain strategy | size-exception |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Low

### Suggested Work Units
| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Complete refactoring and verify | PR 1 | Base branch |

## Phase 1: Infrastructure / Foundation
- [x] 1.1 Remove `article_type` parameter from signature of `QualityAnalyzer.analyze` in [quality_analyzer.py](file:///E:/Python/silvina-editorial/src/domain/quality/quality_analyzer.py).

## Phase 2: Implementation / Wiring
- [x] 2.1 Remove `article_type` from `_quality_analyzer.analyze` invocation in [analyze_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py).

## Phase 3: Testing / Verification
- [x] 3.1 Remove `article_type=None` arguments from `analyze` calls in [test_quality_analyzer.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/test_quality_analyzer.py).
- [x] 3.2 Run pytest to execute domain service and application use case tests and verify all pass.
</Tasks: Remove Unused Article Type Parameter>
