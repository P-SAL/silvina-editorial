# Proposal: Centralize Configuration

## Intent
Eliminate technical debt from scattered `os.getenv` and `os.environ` reads across wirings and adapters by centralizing environment variable configuration in a single class `EnvConfig`.

> [!NOTE]
> While Clean Architecture conventions prefer `EnvironmentConfig` to avoid abbreviations, the user explicitly chose `EnvConfig` in `env_config.py` for local consistency.

## Scope

### In Scope
- Create `src/infrastructure/env_config.py` parsing env variables at instantiation and caching them as instance attributes.
- Refactor `AnalyzeDocumentUseCaseWiring` and `ExportReportWiring` to instantiate `EnvConfig` and inject configuration.
- Inject env values into `DocxReportSettings` constructor, removing dynamic env reads from it.
- Delete `src/infrastructure/config/recommendation_config.py`.
- Add all environment variables with defaults to `.env` and `.env.example`.
- Create `src/infrastructure/tests/test_env_config.py`.

### Out of Scope
- Changing domain models or business rules.
- Modifying UI logic or CLI argument parsing.

## Capabilities

### New Capabilities
None

### Modified Capabilities
- `analyze-document`: `RecommendationConfig` is replaced by `EnvConfig` for recommendation settings.
- `export-report`: Environment configuration for report settings is injected via wiring using `EnvConfig`.

## Approach
1. **Define `EnvConfig`** (`src/infrastructure/env_config.py`): Parse env vars in `__init__` with defaults, caching them as typed instance attributes. Provide `get_recommendation_settings() -> RecommendationSettingsDTO`.
2. **Refactor `DocxReportSettings`**: Remove `default_factory=lambda: environ.get(...)` and keep static defaults.
3. **Refactor Wirings**: Update `AnalyzeDocumentUseCaseWiring` and `ExportReportWiring` to instantiate `EnvConfig` and inject settings.
4. **Delete `recommendation_config.py`**.
5. **Update `.env`/`.env.example`** with all environment variables.
6. **Update/Create Tests**: Refactor existing wiring/settings tests and add new unit tests in `src/infrastructure/tests/test_env_config.py`.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/infrastructure/env_config.py` | New | Central environment config class. |
| `src/infrastructure/tests/test_env_config.py` | New | Unit tests for configuration parsing. |
| `src/infrastructure/config/recommendation_config.py` | Removed | Configuration logic moved to `EnvConfig`. |
| `src/infrastructure/wirings/analyze_document_use_case_wiring.py` | Modified | Uses `EnvConfig` to construct dependencies. |
| `src/infrastructure/wirings/export_report_wiring.py` | Modified | Uses `EnvConfig` to inject report settings. |
| `src/infrastructure/adapters/report/docx_report_settings.py` | Modified | Removes direct env reads; receives injected values. |
| `.env` / `.env.example` | Modified | Declares all environment variables. |
| `src/infrastructure/tests/` | Modified | Updates wiring/settings tests. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Missing env variable in test suite | Low | Set default values in `EnvConfig` for all keys. |
| Test environment isolation issues | Med | Instantiate `EnvConfig` per-run / wiring invocation. |

## Rollback Plan
Revert changes using git:
```bash
git checkout HEAD -- src/
git checkout HEAD -- .env .env.example
```

## Dependencies
None

## Success Criteria
- [ ] All environment variable reads are isolated within `EnvConfig`.
- [ ] No wiring or adapter uses `os.getenv` or `os.environ` directly.
- [ ] `test_env_config.py` asserts parsing, type casting, defaults, and DTO builders.
- [ ] All unit tests pass successfully.
