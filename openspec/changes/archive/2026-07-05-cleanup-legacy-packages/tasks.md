# Tasks: Cleanup Legacy Packages

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 2500 - 3500 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 (Adapt Tests) → PR 2 (Delete Legacy Code) |
| Delivery strategy | ask-on-risk |
| Chain strategy | feature-branch-chain |

Decision needed before apply: No
Chained PRs recommended: Yes
Chain strategy: feature-branch-chain
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Adapt tests to use `src/` directly and remove legacy mocks. | PR 1 | Base branch: main. Verify all tests pass. |
| 2 | Delete legacy packages, root files, and legacy tests. | PR 2 | Base branch: PR 1 branch. Verify all tests pass. |

## Phase 1: Adapt Active Tests (PR 1)

- [x] 1.1 In [test_cli_e2e.py](file:///E:/Python/silvina-editorial/tests/e2e/test_cli_e2e.py), remove legacy `data_access.word_counter` patches.
- [x] 1.2 In [test_gradio_e2e.py](file:///E:/Python/silvina-editorial/tests/e2e/test_gradio_e2e.py), remove legacy `data_access.word_counter` patches.
- [x] 1.3 In [test_read_document_parity.py](file:///E:/Python/silvina-editorial/tests/smoke/test_read_document_parity.py), change assertions to check `src` adapter output directly and remove legacy reader import.
- [x] 1.4 In [test_extract_content_parity.py](file:///E:/Python/silvina-editorial/tests/smoke/test_extract_content_parity.py), change assertions to check `src` adapter output directly and remove legacy extractor import.
- [x] 1.5 In [test_validate_structure_parity.py](file:///E:/Python/silvina-editorial/tests/smoke/test_validate_structure_parity.py), change assertions to validate against `src` validator directly and remove legacy validator/enums imports.
- [x] 1.6 In [test_classify_article_parity.py](file:///E:/Python/silvina-editorial/tests/smoke/test_classify_article_parity.py), change assertions to test `src` classifier directly with mock Client and remove legacy classifier/imports.
- [x] 1.7 In [pytest.ini](file:///E:/Python/silvina-editorial/pytest.ini), remove `tests/legacy` from `norecursedirs`.
- [x] 1.8 Run `pytest` to verify adapted tests pass with legacy files still present.

## Phase 2: Delete Legacy Files & Folders (PR 2)

- [x] 2.1 Delete legacy source directories: `domain/`, `data_access/`, `business_logic/`, `presentation/`.
- [x] 2.2 Delete legacy root files: `apa_validator.py`, `eumic_verifier.py`, `config.py`, `main_legacy.py`.
- [x] 2.3 Delete legacy tests directory: `tests/legacy/`.
- [x] 2.4 Run `pytest` to verify the codebase and adapted test suite function completely without legacy code.
