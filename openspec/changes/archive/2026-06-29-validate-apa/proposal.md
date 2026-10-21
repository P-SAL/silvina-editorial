# Proposal: validate-apa (Slice 3)

## Intent

`apa_validator.py` (root) is a fully pure computation module with no external dependencies,
currently accessed via a module-level function `validate_apa_citations()` that mixes domain
logic with presentation concerns (`generate_report()`). The hexagonal migration plan (§4.2–4.4)
requires this module to be lifted into the domain layer as a stateless service with a proper use
case and wiring, eliminating the root-level convenience function and the dead `self.violations`
state. This unblocks downstream slices that depend on `ApaViolation` and `ApaValidationResult`
DTOs being available as proper domain types.

## Scope

### In Scope

- `src/domain/enums/apa_error_type.py` — migrate `APAErrorType` enum (keep all 8 values, including 2 unused)
- `src/domain/dtos/apa_violation_dto.py` — frozen DTO replacing `APAViolation` dataclass
- `src/domain/dtos/apa_validation_result_dto.py` — new frozen DTO: `is_valid`, `violation_count`, `violations`
- `src/domain/citation/apa_validator.py` — stateless `ApaValidator` domain service (drops `self.violations`, drops `generate_report()`)
- `src/application/validate_apa_use_case.py` — `ValidateApaUseCase.execute(citations: list[tuple[str, int, str]]) -> ApaValidationResult`
- `src/infrastructure/wirings/validate_apa_wiring.py` — `ValidateApaWiring` with `_get_*` per dependency pattern
- `src/domain/tests/citation/test_apa_validator.py` — unit tests covering all 9 validation checks

### Out of Scope

- `generate_report()` — presentation concern; deferred to a future formatter adapter (Slice 13 or later)
- Wiring `ValidateApaUseCase` into `main.py` — deferred to Slice 14 (caller switchover)
- Deleting `apa_validator.py` (root) — coexistence maintained until Slice 14
- Application-layer tests for `ValidateApaUseCase` — pure pass-through; domain tests sufficient

## Capabilities

### New Capabilities

- `validate-apa`: APA citation validation as a domain service — stateless computation of
  parenthetical and narrative citation violations, exposed via a use case returning a typed result DTO

### Modified Capabilities

None

## Approach

Follow the Slice 2 (`validate-structure`) pattern exactly:

1. Create `src/domain/citation/` folder with `__init__.py`.
2. Migrate enum → DTO × 2 → domain service in dependency order.
3. `ApaValidator` is stateless: each `validate_citation()` call returns a new list; `validate_all_citations()` flattens them. No `self.violations`.
4. `ValidateApaUseCase.execute()` receives the same 3-tuple list the legacy `main.py` already builds, calls the domain service, and returns `ApaValidationResult`.
5. `ValidateApaWiring.create_use_case()` constructs the use case via `_get_*` accessors (same pattern as `ValidateStructureWiring`).
6. Tests rewrite `tests/test_apa_validator.py` as `unittest.TestCase` under `src/domain/tests/citation/`.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/domain/enums/apa_error_type.py` | New | `ApaErrorType` enum (8 values) |
| `src/domain/dtos/apa_violation_dto.py` | New | Frozen `ApaViolation` DTO |
| `src/domain/dtos/apa_validation_result_dto.py` | New | Frozen `ApaValidationResult` DTO |
| `src/domain/citation/__init__.py` | New | Package init for citation domain |
| `src/domain/citation/apa_validator.py` | New | Stateless `ApaValidator` service |
| `src/application/validate_apa_use_case.py` | New | `ValidateApaUseCase` |
| `src/infrastructure/wirings/validate_apa_wiring.py` | New | `ValidateApaWiring` |
| `src/domain/tests/citation/test_apa_validator.py` | New | Domain service unit tests |
| `apa_validator.py` (root) | Unchanged | Legacy stays alive during coexistence |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Non-author skip patterns (institutional acronyms, arXiv, DOI, date ranges) are business rules that must be preserved exactly | Med | Copy regex/logic verbatim; test each pattern |
| `YEAR_FORMAT_ERROR` / `PARENTHESES_ERROR` unused enum values may be dropped inadvertently | Low | Explicitly document carry-forward in enum file |
| `generate_report()` boundary — if downstream slices expect it from domain layer | Low | Excluded by design; report generation deferred; document decision |

## Rollback Plan

All new files are additive. Legacy `apa_validator.py` (root) is untouched. To roll back: delete
the 7 new source files and the test file. No existing behavior changes. `main.py` continues
importing from root. No migration state to undo.

## Dependencies

- Slice 2 (`validate-structure`) archived — establishes wiring pattern to follow
- `src/domain/exceptions/citation_errors.py` — available (Slice 1), not needed here (pure computation returns empty lists, not exceptions)

## Success Criteria

- [ ] `ApaErrorType` enum has all 8 values, including the 2 currently unused ones
- [ ] `ApaViolation` is a `frozen=True` dataclass DTO with all 6 fields from legacy
- [ ] `ApaValidationResult` is a `frozen=True` dataclass DTO with `is_valid`, `violation_count`, `violations`
- [ ] `ApaValidator.validate_citation()` and `validate_all_citations()` have no `self.violations` state
- [ ] `generate_report()` is absent from the domain service
- [ ] `ValidateApaUseCase.execute(citations: list[tuple[str, int, str]]) -> ApaValidationResult` works end-to-end
- [ ] All 9 validation checks (6 parenthetical + 3 narrative) pass their tests
- [ ] Non-author skip patterns are preserved exactly (institutional acronyms, arXiv, DOI, date ranges)
- [ ] Legacy `apa_validator.py` (root) is unmodified; `main.py` still imports from root
