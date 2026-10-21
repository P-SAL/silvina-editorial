# Proposal: Load version from file

## Intent
Decouple software versioning from environment variables by removing `SILVINA_VERSION` from `.env` and loading the version dynamically from a root `version.txt` file.

## Scope

### In Scope
- Create root `version.txt` with plain string `0.95`.
- Remove `SILVINA_VERSION` from `.env` and `.env.example`.
- Modify `EnvConfig` (`src/infrastructure/env_config.py`) to read `version.txt` resolved relative to itself (`Path(__file__).resolve().parents[2] / "version.txt"`).
- Fail-fast by raising `FileNotFoundError` (or standard OS/permission errors) if the file is missing/unreadable.
- Exemption: If `TESTING` env var is `"True"`/`"true"`/`"1"`, fallback to `SILVINA_VERSION` env var (default `"0.9"`).
- Update unit tests in `test_env_config.py` and `test_export_report_wiring.py`.
- Update config table in `openspec/specs/analyze-document/spec.md`.
- Mark item 7 in `openspec/TECHNICAL_DEBT.md` as resolved.

### Out of Scope
- Updating version references in general documentation (e.g., `README.md`).

## Capabilities

### New Capabilities
None

### Modified Capabilities
- `analyze-document`: `SILVINA_VERSION` is loaded from `version.txt` instead of env config.

## Approach
1. **Create `version.txt`**: Write `0.95` at root.
2. **Remove env var**: Delete `SILVINA_VERSION` from `.env`/`.env.example`.
3. **Refactor `EnvConfig`**:
   - Check `getenv("TESTING", "").lower() in ("true", "1")`. If so, load `silvina_version` from env (default `"0.9"`).
   - Otherwise, if `version.txt` is missing, raise `FileNotFoundError`. Read, strip, and assign version. Let other IO errors propagate.
4. **Update tests**: Use `TESTING=True` or mock filesystem to test fallback and exception raising.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `version.txt` | New | Holds version string. |
| `.env`, `.env.example` | Modified | Removed `SILVINA_VERSION`. |
| `src/infrastructure/env_config.py` | Modified | Load version from file; fail-fast unless `TESTING=True`. |
| `src/infrastructure/tests/*` | Modified | Adjusted to mock file read / handle fallback. |
| `openspec/specs/analyze-document/spec.md` | Modified | Updated configuration table. |
| `openspec/TECHNICAL_DEBT.md` | Modified | Mark item 7 resolved. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Test suite breakage | High | Check `TESTING` env var for fallback behavior in tests. |

## Rollback Plan
Restore `SILVINA_VERSION` in env files, delete `version.txt`, and revert `env_config.py`.

## Dependencies
- None

## Success Criteria
- [ ] `version.txt` exists with content `0.95`.
- [ ] Absence of `version.txt` raises `FileNotFoundError` in production.
- [ ] All tests pass using the testing fallback or mocks.
