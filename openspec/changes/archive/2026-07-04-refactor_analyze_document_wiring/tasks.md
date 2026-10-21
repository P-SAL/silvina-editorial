# Tasks: Refactor Document Analysis Wiring

## Review Workload Forecast

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: stacked-to-main
400-line budget risk: High

| Field | Value |
|-------|-------|
| Estimated changed lines | ~1600 (approx. 200 added, 1400 deleted) |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | Stacked PRs (PR 1: Orchestrator & Wiring, PR 2: Deletions & Tests) |
| Delivery strategy | ask-on-risk |
| Chain strategy | stacked-to-main |

### Suggested Work Units
| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| Unit 1 | Refactor Orchestrator & Wiring | PR 1 | Main refactoring logic. |
| Unit 2 | Update orchestrator tests | PR 1 | Verify new signature. |
| Unit 3 | Delete obsolete files | PR 2 | Bulk deletions to avoid noise. |

## Phase 1: Orchestrator & Wiring Refactoring
- [x] 1.1 Update constructor/execute in [analyze_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py) to directly orchestrate 13 dependencies.
- [x] 1.2 Refactor [analyze_document_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py) to instantiate and inject all 13 dependencies.

## Phase 2: Test Updates & Cleanup
- [x] 2.1 Refactor [test_analyze_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/tests/test_analyze_document_use_case.py) to mock 13 direct dependencies.
- [x] 2.2 Refactor [test_analyze_document_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_analyze_document_use_case_wiring.py) to assert direct fields.
- [x] 2.3 Delete 10 sub-use case files, 10 sub-wiring files, and 20 corresponding test files.
- [x] 2.4 Run pytest to verify all remaining test suites pass.
