# Proposal: Domain Exceptions (Slice 1)

> Slice 1 of the hexagonal migration plan (`docs/plan-migracion-hexagonal.md`).
> Normative guide: `.agent/skills/clean-architecture/SKILL.md`.

## Intent

Populate the **per-group domain exception files** under `src/domain/exceptions/`
so that every domain grouping that needs to raise errors has a named, typed
exception class instead of relying on `ValueError` or `Exception` generic raises.

This slice gives later slices (ValidateStructure, ValidateApa, adapters, use
cases) a clean vocabulary to raise against. Slice 2 onward raise these specific
`BaseSrcError` subclasses; until this slice lands they have nowhere clean to
raise to. This slice unblocks all error-raising migration in slices 2–16.

> **Coexistence**: the legacy code (`business_logic/`, `data_access/`, entry
> points) continues to use `ValueError`, `Exception`, and `try/except` with
> `print`. No legacy file is touched. Exception files are **defined here** but
> not yet raised by any use case — each later slice wires them in as it
> migrates its own logic.

### Why now

The migration plan places exception population at Slice 1 — immediately after
the domain foundations (Slice 0) — because:

1. The base hierarchy (`BaseSrcError`, `SrcBaseWarning`, `SrcBaseNotFound`) is
   already in `base_src_error.py` (Slice 0 prerequisite; it was in the skeleton
   from the start).
2. Defining exceptions before use cases prevents slices from choosing arbitrary
   base types on the fly.
3. The domain exception files have **zero external dependencies** — pure Python,
   pure domain, no I/O. They are the safest possible next step.

### Success looks like

- Five exception group files exist under `src/domain/exceptions/`, each holding
  the exception classes called for by `docs/plan-migracion-hexagonal.md` §7.
- Each class inherits the correct base type (`SrcBaseNotFound` or
  `SrcBaseWarning`) and is catchable as `BaseSrcError`.
- Each class defines a `MESSAGE` string attribute.
- Each group file has a matching `unittest.TestCase` in
  `src/domain/tests/exceptions/`.
- `python -m pytest src/` passes green (currently 120 tests; this slice adds
  the 5 new exception group test files).
- No legacy file is modified.

## Scope

### In Scope

- **5 exception group files** under `src/domain/exceptions/`:
  - `document_errors.py` — `DocumentNotFound`, `DocumentEmpty`,
    `DocumentUnreadable`
  - `citation_errors.py` — `CitationParsingFailed`
  - `classification_errors.py` — `ClassificationFailed`
  - `quality_errors.py` — `QualityAnalysisFailed`
  - `language_model_errors.py` — `LanguageModelUnavailable`
- **5 test files** under `src/domain/tests/exceptions/`.
- Coexistence with all legacy code and the existing `src/` test suite.

### Out of Scope

- Actually raising these exceptions in any use case, adapter, or service —
  each later slice does that when it migrates its own logic.
- Adding new exception groups beyond the 5 defined in the plan §7.
- Modifying `base_src_error.py` or the `generic_error_handler` decorator.
- Any I/O, ports, adapters, or wirings.
- Deleting or modifying any legacy file.

## Approach

One exception group file per domain grouping. Each file may hold multiple
closely related exception classes — this is the documented exception to the
one-class-per-file rule (SKILL §4 + §5). Each class inherits the correct base
type and defines a `MESSAGE` class attribute. Tests verify inheritance and
catchability as `BaseSrcError`.

**Per-group loop (Strict TDD, runner `python -m pytest src/`):**

1. Write failing `unittest.TestCase` in
   `src/domain/tests/exceptions/test_<group>_errors.py`.
2. Create `src/domain/exceptions/<group>_errors.py` with the exception classes.
3. Run `python -m pytest src/` green; move to the next group.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/domain/exceptions/` | 5 new files | One per domain grouping |
| `src/domain/tests/exceptions/` | 5 new files | `unittest.TestCase` per group |
| `base_src_error.py` | Untouched | Base hierarchy already in place |
| All legacy code | Untouched | Coexistence; no caller is rewired |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Choosing wrong base type for an exception | Low | Plan §7 specifies each; verified by tests |
| Future slice adding exceptions to wrong file | Low | Each group file's name matches its domain folder |
| Slice adds boilerplate with no immediate caller | Low | Intentional — this is a foundation slice; plan §8 documents this pattern |

## Rollback Plan

All work is additive under `src/domain/exceptions/` and `src/domain/tests/`.
No production code raises these exceptions yet. Rollback = delete the 5 new
group files and their tests; no behavior changes.

## Dependencies

- Python 3.10+.
- `src/domain/exceptions/base_src_error.py` — already present (skeleton).
- No external libraries (pure domain).

## Success Criteria

- [ ] All 5 exception group files exist with the classes listed in plan §7.
- [ ] Each exception inherits the correct base type.
- [ ] Each exception has a `MESSAGE` class attribute.
- [ ] `python -m pytest src/` green (≥120 + new tests).
- [ ] No legacy file is modified.
