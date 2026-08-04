# Proposal: Resolve Pytest Domain Namespace Collision

## Intent

Resolve [TECHNICAL_DEBT.md Item 1](file:///E:/Python/silvina-editorial/openspec/TECHNICAL_DEBT.md#L9-L15) to fix the bare `pytest` collection failure caused by a namespace collision between legacy `domain/` and new `src/domain/`.

## Scope

### In Scope
- Add `--import-mode=importlib` to the `pytest` configuration section in `pytest.ini`.

### Out of Scope
- Renaming the legacy `domain/` directory at the repository root.
- Renaming the `src/domain/` directory.

## Capabilities

### New Capabilities
None

### Modified Capabilities
None

## Approach

Use modern pytest best practice by setting `addopts = --import-mode=importlib` in `pytest.ini`. This forces pytest to import test modules directly from their file paths without adding parent directories to `sys.path`, avoiding namespace collision and resolving `ModuleNotFoundError`.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `pytest.ini` | Modified | Add `addopts = --import-mode=importlib` under `[pytest]` |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Test/fixture relies on `sys.path` prepend side effects | Low | Codebase already uses explicit `src.` prefix imports |

## Rollback Plan

Revert the change in `pytest.ini` by removing `--import-mode=importlib` (restoring it to the original file content).

## Dependencies

None

## Success Criteria

- [ ] Running `pytest` from the repository root without path scoping collects all tests (currently 287 tests) without any collection errors (resolving the 104 current collection errors).
