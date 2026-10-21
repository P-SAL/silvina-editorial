# Proposal: Cleanup Legacy Packages

## Intent
Remove unused legacy root packages and files that have been superseded by the new hexagonal architecture under `src/`. Adapt `tests/smoke/` integration/parity tests to run directly against the new architecture classes to preserve regression test coverage.

## Scope

### In Scope
- Delete legacy packages: `domain/`, `data_access/`, `business_logic/`, `presentation/`.
- Delete legacy root files: `apa_validator.py`, `eumic_verifier.py`, `config.py`, `main_legacy.py`.
- Delete legacy test folder: `tests/legacy/`.
- Adapt `tests/smoke/` tests (`test_classify_article_parity.py`, `test_extract_content_parity.py`, `test_read_document_parity.py`, `test_validate_structure_parity.py`) to test `src/` classes directly by removing legacy imports.
- Update `tests/e2e/test_cli_e2e.py` and `tests/e2e/test_gradio_e2e.py` to remove legacy patches (`data_access.word_counter.WIN32COM_AVAILABLE`).
- Update `pytest.ini` to remove the `tests/legacy` reference from `norecursedirs`.

### Out of Scope
- Modifying the core business logic of any classes under `src/`.

## Capabilities

### New Capabilities
- None

### Modified Capabilities
- None

## Approach
1. Delete all legacy directories (`domain/`, `data_access/`, `business_logic/`, `presentation/`, `tests/legacy/`) and legacy root files (`apa_validator.py`, `eumic_verifier.py`, `config.py`, `main_legacy.py`).
2. Update E2E test files (`tests/e2e/test_cli_e2e.py`, `tests/e2e/test_gradio_e2e.py`) to remove the patch on `data_access.word_counter.WIN32COM_AVAILABLE`. Keep the patch on `src.infrastructure.adapters.document.win32com_word_count_adapter.WIN32COM_AVAILABLE`.
3. In `tests/smoke/`, replace legacy imports with the corresponding `src/` classes and verify their integration.
4. Modify `pytest.ini` to remove `tests/legacy` from `norecursedirs`.
5. Run the test suite to verify everything passes.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `domain/`, `data_access/`, `business_logic/`, `presentation/` | Removed | Legacy source packages |
| `apa_validator.py`, `eumic_verifier.py`, `config.py`, `main_legacy.py` | Removed | Legacy entry points & config |
| `tests/legacy/` | Removed | Legacy test suites |
| `tests/smoke/` | Modified | Adapt smoke tests to target `src/` directly |
| `tests/e2e/test_cli_e2e.py`, `tests/e2e/test_gradio_e2e.py` | Modified | Clean up legacy mocks/patches |
| `pytest.ini` | Modified | Remove legacy path from test exclusion list |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| E2E tests fail due to missing modules when patching `data_access.word_counter` | High | Remove the legacy patch and ensure the new win32com adapter patch remains in place |
| Smoke tests fail due to interface mismatches between legacy and new classes | Low | Adjust smoke tests to use new DTOs/Entities and Use Cases |

## Rollback Plan
Discard working changes via git checkout/clean:
```bash
git checkout -- pytest.ini tests/e2e/ tests/smoke/
git clean -fd domain/ data_access/ business_logic/ presentation/ tests/legacy/ apa_validator.py eumic_verifier.py config.py main_legacy.py
```

## Dependencies
- None

## Success Criteria
- [ ] All tests run and pass using `.venv/Scripts/pytest`.
- [ ] All legacy root packages and files are deleted.
- [ ] Adapted smoke tests verify `src/` implementations successfully.
