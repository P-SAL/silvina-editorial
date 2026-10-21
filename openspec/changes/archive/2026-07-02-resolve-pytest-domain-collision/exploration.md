## Exploration: pytest bare collection collision

### Current State
Running `pytest` from the repository root (without path scoping) fails to collect tests under `src/domain/tests/`, throwing `ModuleNotFoundError: No module named 'domain.tests'`.

This is caused by a package name collision between the legacy `domain/` directory at the repository root and the new `src/domain/` directory. When pytest discovers tests in both the root `tests/` directory and `src/domain/tests/`, it adds both the repository root (`E:\Python\silvina-editorial`) and `src/` to `sys.path`.

Since `src/` is on `sys.path`, the test package is imported as `domain.tests...`. However, because the repository root is also on `sys.path` (and typically checked first or resolved preferentially), Python attempts to resolve `domain` to the legacy `domain/` directory at the root. Since the legacy `domain/` package has no `tests` subdirectory, Python raises `ModuleNotFoundError`.

### Affected Areas
- `pytest.ini` — Needs configuration changes to change how test modules are imported.
- `domain/` — Legacy package at the repository root colliding with `src/domain/`.
- `src/domain/tests/` — Test package whose discovery triggers the collision.

### Approaches
1. **Configure `--import-mode=importlib` in `pytest.ini`** — Tell pytest to import test modules directly from their file paths without adding their parent directories to `sys.path`. The modules are imported with names like `src.domain.tests.exceptions.test_base_src_error`.
   - Pros:
     - Simple 1-line configuration change.
     - Avoids namespace pollution and import collisions.
     - Modern pytest best practice (default in pytest 8+).
     - Keeps legacy `domain/` intact as required by project migration plans.
   - Cons: None.
   - Effort: Low

2. **Rename legacy `domain/` package** — Rename the legacy `domain/` folder at the root to `domain_legacy/` or `legacy_domain/`.
   - Pros:
     - Eliminates the duplicate name from the import search paths.
   - Cons:
     - Requires renaming the folder and auditing/updating imports in legacy code.
     - Violates the decision to defer legacy package deletion/cleanup to Slice 16 (cleanup).
   - Effort: Medium

3. **Rename `src/domain/` package** — Rename the new domain package under `src/`.
   - Pros:
     - Resolves name collision.
   - Cons:
     - Breaks standard clean architecture / DDD conventions.
     - Requires updating all imports across the entire migrated codebase (huge churn).
   - Effort: High

### Recommendation
We recommend **Approach 1 (Configure `--import-mode=importlib` in `pytest.ini`)**. It is the cleanest, least intrusive solution, respects the project's migration constraints regarding legacy packages, and aligns with modern pytest standards.

### Risks
- If any tests or fixtures rely on `sys.path` modification side-effects of pytest's default `prepend` import mode, they might fail. However, since the codebase imports all migrated modules via the `src.` prefix, this risk is extremely low.

### Ready for Proposal
Yes — The root cause is fully understood, and the fix is verified to be a simple, standard configuration adjustment. The orchestrator can proceed to the proposal phase.
