# Task Checklist: Domain Exceptions (Slice 1)

> Generated from: spec `sdd/domain-exceptions/spec` + design `sdd/domain-exceptions/design`
> Runner: `python -m pytest src/`
> TDD order: failing-first per group (red → green per work unit)
> Coexistence rule: no legacy file is modified; all tasks are purely additive.
> Status: ALL TASKS COMPLETE — code is implemented and green (120 src tests pass).

---

## Review Workload Forecast

| Metric | Estimate |
|---|---|
| Production files created | 5 (one per exception group) |
| Test files created | 5 |
| Estimated lines — production (~8–20 lines/file) | ~55 |
| Estimated lines — tests (~15 lines/file) | ~75 |
| **Total estimated changed lines** | **~130** |
| 400-line budget risk | **Low** |
| Chained PRs recommended | No — fits comfortably in one PR |

---

## Prerequisites (verified)

- [x] `src/domain/exceptions/base_src_error.py` exists with `BaseSrcError`,
      `SrcBaseWarning`, `SrcBaseNotFound`, `SrcBaseGenericError`
- [x] `src/domain/tests/exceptions/__init__.py` exists
- [x] `python -m pytest src/` green before this slice (120 tests)

---

## Task 1 — `document_errors.py` + test

**Spec**: REQ-EXC-DOC-1, REQ-EXC-DOC-2, REQ-EXC-DOC-3, REQ-EXC-MSG-1
**Parallel**: Yes (with Tasks 2–5; all groups are independent)

- [x] Write failing test `src/domain/tests/exceptions/test_document_errors.py`
  - `TestDocumentNotFound.test_is_subclass_of_src_base_not_found`
  - `TestDocumentNotFound.test_is_catchable_as_base_src_error`
  - `TestDocumentEmpty.test_is_subclass_of_src_base_warning`
  - `TestDocumentEmpty.test_is_catchable_as_base_src_error`
  - `TestDocumentUnreadable.test_is_subclass_of_src_base_warning`
  - `TestDocumentUnreadable.test_is_catchable_as_base_src_error`
- [x] Create `src/domain/exceptions/document_errors.py`
  - Imports: `from src.domain.exceptions.base_src_error import SrcBaseNotFound, SrcBaseWarning`
  - `DocumentNotFound(SrcBaseNotFound)` — `MESSAGE = "The document file could not be found."`
  - `DocumentEmpty(SrcBaseWarning)` — `MESSAGE = "The document has no readable content."`
  - `DocumentUnreadable(SrcBaseWarning)` — `MESSAGE = "The document could not be read."`
- [x] Run `python -m pytest src/domain/tests/exceptions/test_document_errors.py` — green

**Work unit commit**: `feat(domain/exceptions): add document exception group and tests`

---

## Task 2 — `citation_errors.py` + test

**Spec**: REQ-EXC-CIT-1, REQ-EXC-MSG-1
**Parallel**: Yes (with Tasks 1, 3–5)

- [x] Write failing test `src/domain/tests/exceptions/test_citation_errors.py`
  - `TestCitationParsingFailed.test_is_subclass_of_src_base_warning`
  - `TestCitationParsingFailed.test_is_catchable_as_base_src_error`
- [x] Create `src/domain/exceptions/citation_errors.py`
  - Imports: `from src.domain.exceptions.base_src_error import SrcBaseWarning`
  - `CitationParsingFailed(SrcBaseWarning)` — `MESSAGE = "The citation could not be parsed."`
- [x] Run `python -m pytest src/domain/tests/exceptions/test_citation_errors.py` — green

**Work unit commit**: `feat(domain/exceptions): add citation exception group and tests`

---

## Task 3 — `classification_errors.py` + test

**Spec**: REQ-EXC-CLASS-1, REQ-EXC-MSG-1
**Parallel**: Yes (with Tasks 1–2, 4–5)

- [x] Write failing test `src/domain/tests/exceptions/test_classification_errors.py`
  - `TestClassificationFailed.test_is_subclass_of_src_base_warning`
  - `TestClassificationFailed.test_is_catchable_as_base_src_error`
- [x] Create `src/domain/exceptions/classification_errors.py`
  - Imports: `from src.domain.exceptions.base_src_error import SrcBaseWarning`
  - `ClassificationFailed(SrcBaseWarning)` — `MESSAGE = "The article classification could not be completed."`
- [x] Run `python -m pytest src/domain/tests/exceptions/test_classification_errors.py` — green

**Work unit commit**: `feat(domain/exceptions): add classification exception group and tests`

---

## Task 4 — `quality_errors.py` + test

**Spec**: REQ-EXC-QUAL-1, REQ-EXC-MSG-1
**Parallel**: Yes (with Tasks 1–3, 5)

- [x] Write failing test `src/domain/tests/exceptions/test_quality_errors.py`
  - `TestQualityAnalysisFailed.test_is_subclass_of_src_base_warning`
  - `TestQualityAnalysisFailed.test_is_catchable_as_base_src_error`
- [x] Create `src/domain/exceptions/quality_errors.py`
  - Imports: `from src.domain.exceptions.base_src_error import SrcBaseWarning`
  - `QualityAnalysisFailed(SrcBaseWarning)` — `MESSAGE = "The quality analysis could not be completed."`
- [x] Run `python -m pytest src/domain/tests/exceptions/test_quality_errors.py` — green

**Work unit commit**: `feat(domain/exceptions): add quality exception group and tests`

---

## Task 5 — `language_model_errors.py` + test

**Spec**: REQ-EXC-LM-1, REQ-EXC-MSG-1
**Parallel**: Yes (with Tasks 1–4)

- [x] Write failing test `src/domain/tests/exceptions/test_language_model_errors.py`
  - `TestLanguageModelUnavailable.test_is_subclass_of_src_base_warning`
  - `TestLanguageModelUnavailable.test_is_catchable_as_base_src_error`
- [x] Create `src/domain/exceptions/language_model_errors.py`
  - Imports: `from src.domain.exceptions.base_src_error import SrcBaseWarning`
  - `LanguageModelUnavailable(SrcBaseWarning)` — `MESSAGE = "The language model backend is unavailable."`
- [x] Run `python -m pytest src/domain/tests/exceptions/test_language_model_errors.py` — green

**Work unit commit**: `feat(domain/exceptions): add language model exception group and tests`

---

## Integration Verification

- [x] Run `python -m pytest src/` — full suite green (120 tests; all 5 new group test files included)
- [x] Coexistence confirmed — no legacy file modified

---

## Task Dependency Graph

```
Tasks 1–5 are fully parallel (no intra-slice dependencies)

  Task 1 (document_errors)
  Task 2 (citation_errors)        ← all parallel
  Task 3 (classification_errors)
  Task 4 (quality_errors)
  Task 5 (language_model_errors)
        │
        └── Integration verification (after all 5 tasks green)
```

---

## Summary

| Phase | Tasks | Sequential? | Estimated lines |
|---|---|---|---|
| Exception groups | T1–T5 (all parallel) | No | ~130 |
| **Total** | **5 tasks** | — | **~130** |
