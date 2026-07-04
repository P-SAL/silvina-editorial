# Tasks: Resolve Pytest Domain Namespace Collision

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 2 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | single-pr |
| Chain strategy | pending |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: pending
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Configure pytest and verify execution | PR 1 | Base branch: main; tests/verification included |

## Phase 1: Implementation

- [x] 1.1 Add `addopts = --import-mode=importlib` to [pytest.ini](file:///E:/Python/silvina-editorial/pytest.ini) under the `[pytest]` section. Also added `pythonpath = .` and an empty `src/__init__.py` — both were required in addition to the addopts line; see apply-progress notes.

## Phase 2: Verification

- [x] 2.1 Execute `pytest` from the repository root to confirm all tests collect and run successfully. Result: 635 passed, 3 skipped, 6 subtests passed, 0 collection errors (was 104 collection errors on baseline `addopts`-only attempt, 138 on a first naive attempt).
- [x] 2.2 Run `ruff check` on the codebase to ensure styling and import conventions are clean. Result: 465 pre-existing findings, none in the two files touched by this change (`pytest.ini` is not Python; new `src/__init__.py` is empty and flagged nothing).
