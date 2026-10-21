## Exploration: Slice 16: final cleanup and removal of legacy root packages

### Current State
During the hexagonal migration (Slices 0-15), the codebase has been successfully transitioned to the new Clean Hexagonal Architecture located in `src/`. The legacy packages and modules are no longer used by the new entry points (`main.py` and `gradio_app.py`). However, the legacy packages (`domain/`, `data_access/`, `business_logic/`, `presentation/`), legacy files (`apa_validator.py`, `eumic_verifier.py`, `config.py`, `main_legacy.py`), and corresponding test folders (`tests/legacy/`, `tests/smoke/`) still coexist with the new implementation.

Deleting these files requires resolving a few references inside E2E tests (`tests/e2e/test_cli_e2e.py` and `tests/e2e/test_gradio_e2e.py`) that patch legacy modules (specifically `data_access.word_counter.WIN32COM_AVAILABLE`).

### Affected Areas
The following folders and files will be removed:
- `domain/` (legacy root package) — Fully replaced by `src/domain/`.
- `data_access/` (legacy root package) — Fully replaced by `src/infrastructure/adapters/` and `src/infrastructure/wirings/`.
- `business_logic/` (legacy root package) — Fully replaced by `src/application/` use cases.
- `presentation/` (legacy root package) — Fully replaced by `src/infrastructure/adapters/report/`.
- `apa_validator.py` (legacy root file) — Fully replaced by `src/domain/citation/apa_validator.py`.
- `eumic_verifier.py` (legacy root file) — Fully replaced by `src/application/verify_eumic_use_case.py`.
- `config.py` (legacy root file) — Fully replaced by environment-variable configuration.
- `main_legacy.py` (legacy root file) — Fully replaced by `main.py`.
- `tests/legacy/` (test folder) — Tests for legacy packages.
- `tests/smoke/` (test folder) — Parity tests comparing legacy vs new modules.

The following files will need modification to clean up legacy references:
- `tests/e2e/test_cli_e2e.py` — Remove redundant calls to `patch("data_access.word_counter.WIN32COM_AVAILABLE", False)`.
- `tests/e2e/test_gradio_e2e.py` — Remove redundant calls to `patch("data_access.word_counter.WIN32COM_AVAILABLE", False)` and redirect to `src.infrastructure.adapters.document.win32com_word_count_adapter.WIN32COM_AVAILABLE` where necessary.
- `pytest.ini` — Remove references to `tests/legacy` from `norecursedirs`.

### Approaches

1. **Complete Removal (Recommended)**
   - **Description**: Delete all legacy packages, files, legacy tests, and parity tests. Clean up the references to the legacy modules in E2E tests.
   - **Pros**:
     - Eliminates all dead code.
     - Resolves the top-level namespace collision risk (e.g. `domain` vs `src/domain`).
     - Simplifies the codebase structure and project navigation.
     - Keeps testing clean and aligned with the hexagonal architecture.
   - **Cons**:
     - Parity tests are deleted (but their role is complete as the migration is finished).
   - **Effort**: Low

2. **Selective Preservation**
   - **Description**: Delete legacy source packages but attempt to keep parity smoke tests by mocking or pointing them to `src/`.
   - **Pros**:
     - Retains more smoke test coverage.
     - **Cons**:
     - Keeping parity tests when there is no legacy code left to compare against defeats their purpose and introduces maintenance overhead.
   - **Effort**: Medium

### Recommendation
Proceed with **Approach 1 (Complete Removal)**. The hexagonal migration is fully complete and all 592 test cases pass. Retaining legacy code or parity tests is unnecessary and adds maintenance overhead.

### Risks
- **E2E Test Failures**: If the E2E tests are not updated to remove the patches targeting `data_access.word_counter`, they will fail with `ModuleNotFoundError` during test collection. This is mitigated by updating `tests/e2e/test_cli_e2e.py` and `tests/e2e/test_gradio_e2e.py` alongside the deletions.

### Ready for Proposal
Yes — the project is ready for the proposal phase of `cleanup-legacy-packages`.
