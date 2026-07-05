# Tasks: Resolve Generic Error Handler Scope

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 10-20 |
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
| 1 | Remove generic error handler scope from adapters | PR 1 | Apply changes and verify with the test suite |

## Phase 1: Specification

- [x] 1.1 Modify `openspec/specs/extract-citations/spec.md` to remove the requirement for `@generic_error_handler` on `DocxCitationAdapter` and `DocxReferenceAdapter`.

## Phase 2: Refactoring Adapters

- [x] 2.1 Remove `@generic_error_handler` decorator and its import from `src/infrastructure/adapters/document/docx_text_adapter.py`.
- [x] 2.2 Remove `@generic_error_handler` decorator and its import from `src/infrastructure/adapters/document/docx_citation_adapter.py`.
- [x] 2.3 Remove `@generic_error_handler` decorator and its import from `src/infrastructure/adapters/document/docx_reference_adapter.py`.
- [x] 2.4 Remove `@generic_error_handler` decorator and its import from `src/infrastructure/adapters/document/docx_eumic_adapter.py`.
- [x] 2.5 Remove `@generic_error_handler` decorator and its import from `src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py`.

## Phase 3: Testing & Verification

- [x] 3.1 Run tests for modified document adapters in `src/infrastructure/tests/adapters/document/`.
- [x] 3.2 Run tests for Ollama generator adapter in `src/infrastructure/tests/test_ollama_generator_adapter.py`.
- [x] 3.3 Run full test suite using `.venv/Scripts/pytest` to verify zero regressions.
