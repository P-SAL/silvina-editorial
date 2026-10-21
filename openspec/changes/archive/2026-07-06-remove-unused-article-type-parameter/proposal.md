# Proposal: Remove Unused Article Type Parameter

## Intent
Clean up the codebase by removing the unused `article_type` parameter from `QualityAnalyzer.analyze(...)`.

## Scope
### In Scope
- Remove `article_type` parameter from the signature of `QualityAnalyzer.analyze`.
- Update `AnalyzeDocumentUseCase.execute` call to `_quality_analyzer.analyze`.
- Clean up calls in unit tests in `src/domain/tests/quality/test_quality_analyzer.py`.
### Out of Scope
- Any functional changes to document quality analysis or classification logic.

## Capabilities
### New Capabilities
None
### Modified Capabilities
None

## Approach
- Modify signature of `analyze` in `src/domain/quality/quality_analyzer.py`.
- Remove `article_type` argument from call in `src/application/analyze_document_use_case.py`.
- Remove `article_type=None` keyword argument from calls in `src/domain/tests/quality/test_quality_analyzer.py`.
- Run tests to verify the refactoring.

## Affected Areas
| Area | Impact | Description |
|------|--------|-------------|
| [quality_analyzer.py](file:///E:/Python/silvina-editorial/src/domain/quality/quality_analyzer.py) | Low | Remove parameter from `analyze` signature. |
| [analyze_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py) | Low | Update call site. |
| [test_quality_analyzer.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/test_quality_analyzer.py) | Low | Remove deprecated parameter from test calls. |

## Risks
| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Breaking unknown third-party calls | Very Low | Run global workspace search to ensure all instances are updated. |

## Rollback Plan
- Run `git checkout -- src/` to revert all changes.

## Dependencies
None

## Success Criteria
- [ ] Unit tests in `src/domain/tests/quality/test_quality_analyzer.py` pass.
- [ ] Unit tests in `src/application/tests/test_analyze_document_use_case.py` pass.
- [ ] No occurrences of `article_type` remain in `QualityAnalyzer.analyze` signature or calls.
