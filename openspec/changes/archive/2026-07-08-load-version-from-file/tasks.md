# Tasks: Load version from file

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~70 lines |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | ask-on-risk |
| Chain strategy | stacked-to-main |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: stacked-to-main
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Load version dynamically from root version.txt | PR 1 | Base branch; tests/docs included |

## Phase 1: Environment and File Setup

- [x] 1.1 Create `version.txt` in project root with content `0.95`.
- [x] 1.2 Remove `SILVINA_VERSION` environment variable from `.env`. Done manually by the user (file is outside agent tool access by permission settings).
- [x] 1.3 Remove `SILVINA_VERSION` environment variable from `.env.example`. Confirmed via `git diff` — line removed.

## Phase 2: Configuration and Infrastructure Code

- [x] 2.1 Update `conftest.py` to set `environ["TESTING"] = "True"` at startup.
- [x] 2.2 Refactor `EnvConfig` in `src/infrastructure/env_config.py` to resolve `silvina_version` dynamically.
- [x] 2.3 Implement fallback to `SILVINA_VERSION` env var (default `"0.9"`) in `EnvConfig` if `TESTING` is active.
- [x] 2.4 Raise `FileNotFoundError` in `EnvConfig` when `version.txt` is missing and not in testing mode.

## Phase 3: Tests Verification and Mocking

- [x] 3.1 Update `TestEnvConfig` (`src/infrastructure/tests/test_env_config.py`) to verify `EnvConfig` defaults are loaded when env is empty and `version.txt` exists.
- [x] 3.2 Add tests in `test_env_config.py` to verify `EnvConfig` fails fast when `version.txt` is missing.
- [x] 3.3 Add tests in `test_env_config.py` to verify `EnvConfig` falls back to `SILVINA_VERSION` in testing mode (with/without env override).
- [x] 3.4 Update `test_export_report_wiring.py` to patch `TESTING="True"` during wiring tests.

## Phase 4: Specifications and Tech Debt Cleanup

- [x] 4.1 Update environment variable configuration table in `openspec/specs/analyze-document/spec.md`.
- [x] 4.2 Move technical debt item 7 in `openspec/TECHNICAL_DEBT.md` to the resolved section.
