# Design: Resolve Pytest Domain Namespace Collision

## Technical Approach

Configure `pytest` to use the `importlib` import mode globally by modifying `pytest.ini`. This addresses the namespace collision between the legacy `domain/` directory at the repository root and the new `src/domain/` directory by changing how test modules are loaded, preventing `pytest` from polluting `sys.path` with parent directories during test collection.

## Architecture Decisions

### Decision: Test Module Import Mode

| Option | Tradeoffs | Decision |
|--------|-----------|----------|
| `prepend` | Default pytest behavior. Modifies `sys.path` by prepending each test's parent directory. Causes namespace collisions when two directories contain a folder of the same name (e.g., legacy `domain/` and new `src/domain/`). | Rejected |
| `append` | Appends parent directories to `sys.path`. Does not resolve the collision because both directories are still placed on `sys.path`, leading to ambiguous imports. | Rejected |
| `importlib` | Standard in modern pytest (8.0+). Imports test modules by path without modifying `sys.path` for parent folders. Avoids namespace collision and imports tests under unique path-based names. | **Selected** |

**Rationale**: `importlib` is the modern pytest standard that avoids namespace collisions by not altering `sys.path` for test directories. It enables running `pytest` globally from the repository root without colliding the legacy `domain/` directory with the migrated `src/domain/`.

## Data Flow

No runtime data flow is affected by this change. This is a development and test configuration adjustment.

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `pytest.ini` | Modify | Add `addopts = --import-mode=importlib` and `pythonpath = .` to the `[pytest]` section. |
| `src/__init__.py` | Create | Empty marker file, required alongside `--import-mode=importlib` (see Deviations below). |

## Interfaces / Contracts

No new interfaces or API contracts are introduced. The modification affects the test collection behavior in `pytest.ini`.

```ini
[pytest]
norecursedirs = tests/legacy .venv .git
addopts = --import-mode=importlib
pythonpath = .
```

## Deviations from Initial Design (discovered during apply)

The single `addopts` line was not sufficient in practice; two additional changes were required:

1. **`pythonpath = .`** — Without it, `--import-mode=importlib` alone broke `from src.xxx import ...` in `src/application/tests/*` (`ModuleNotFoundError: No module named 'src'`). Unlike the default `prepend` mode, `importlib` mode does not implicitly add the repository root to `sys.path`.
2. **`src/__init__.py` (new, empty)** — `src/` itself lacked an `__init__.py`. Under `importlib` mode, pytest still inserted a basedir for packaged test files, registering test modules as top-level `domain.tests...` and binding `sys.modules['domain']` to `src/domain`, which shadowed the legacy root `domain/models.py` package that `tests/smoke/*_parity.py` depends on via `business_logic/*`.

Both were verified against the full suite: naive `addopts`-only attempt produced 138 collection errors (worse than the 104 baseline); the combined fix produced 0 collection errors across 635 passing tests. Scope was respected — neither legacy `domain/` nor `src/domain/` were renamed or modified.

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | N/A | No new python code is written. |
| Integration | Test Collection and Execution | Run `pytest` from the repository root. Ensure that all 287+ tests (including legacy and migrated tests) collect and execute without raising namespace `ModuleNotFoundError`. |

## Migration / Rollout

No migration or database rollout required. The change takes effect immediately upon updating the configuration file.

## Open Questions

None. The solution is straightforward and fully mitigates the namespace collision issue.
