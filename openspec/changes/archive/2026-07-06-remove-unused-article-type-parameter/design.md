<Design: Remove Unused Article Type Parameter>
# Design: Remove Unused Article Type Parameter

## Technical Approach
The `article_type` parameter is currently accepted by `QualityAnalyzer.analyze` but is not utilized anywhere in the method's implementation. To simplify the interface and eliminate dead code, we will remove this parameter from the signature and update all callers and unit tests.

A codebase search verified that there are no references to `.analyze(` on a `QualityAnalyzer` instance other than:
- `src/application/analyze_document_use_case.py`
- `src/domain/tests/quality/test_quality_analyzer.py`
- `src/application/tests/test_analyze_document_use_case.py` (which uses a MagicMock, so it will continue working without modification but is verified as part of the use-case test suite).

## Architecture Decisions
| Decision | Choice | Alternatives considered | Rationale |
|---|---|---|---|
| Remove `article_type` parameter | Remove the parameter entirely from the `QualityAnalyzer.analyze` signature. | Keep the parameter as optional with a default value. | The parameter is completely unused; keeping it introduces dead code, noise, and confuses the interface's dependencies. |

## Data Flow
```
[AnalyzeDocumentUseCase]
       │
       ├─► [ArticleClassifier.classify(document_content)]
       │         │
       │         ▼ (returns classification)
       │
       └─► [QualityAnalyzer.analyze(document_content)] (article_type parameter removed)
                 │
                 ▼ (returns QualityResultDTO)
```

## File Changes
| File | Action | Description |
|------|--------|-------------|
| [quality_analyzer.py](file:///E:/Python/silvina-editorial/src/domain/quality/quality_analyzer.py) | Modify | Remove `article_type` from the signature of `analyze`. |
| [analyze_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py) | Modify | Update call to `_quality_analyzer.analyze` to omit `article_type`. |
| [test_quality_analyzer.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/test_quality_analyzer.py) | Modify | Remove `article_type=None` keyword arguments from all `analyze` test calls. |

## Interfaces / Contracts

### Before
```python
class QualityAnalyzer:
    def analyze(self, document_content: DocumentContentDTO, article_type) -> QualityResultDTO:
        ...
```

### After
```python
class QualityAnalyzer:
    def analyze(self, document_content: DocumentContentDTO) -> QualityResultDTO:
        ...
```

## Testing Strategy
| Layer | What to Test | Approach |
|---|---|---|
| Domain Service | `QualityAnalyzer` behavior | Run tests in `test_quality_analyzer.py` to ensure analyzer correctly scores and resolves quality levels without the parameter. |
| Application Use Case | `AnalyzeDocumentUseCase` orchestration | Run tests in `test_analyze_document_use_case.py` to check that the pipeline completes successfully and quality analysis is triggered exactly once. |

## Migration / Rollout
No migration required.

## Open Questions
None
</Design: Remove Unused Article Type Parameter>
