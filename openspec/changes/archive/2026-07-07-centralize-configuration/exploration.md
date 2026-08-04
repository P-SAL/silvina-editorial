## Exploration: Centralize Configuration

### Current State
Currently, environment variables are read in multiple disjointed places across the infrastructure layer:
1. `RecommendationConfig` reads environment variables at call time when building `RecommendationSettingsDTO`.
2. `DocxReportSettings` reads environment variables using dynamic `default_factory` fields (using `os.environ.get` at construction time).
3. `AnalyzeDocumentUseCaseWiring` reads environment variables using `os.getenv` directly within its methods to instantiate adapters and DTOs (e.g., `CITATION_MAX_AUTHOR_NAME_LENGTH`, `GRAMMAR_MAX_REPLACEMENTS`, `STRUCTURE_MAX_HEADER_LENGTH`, classifier temperature, article size thresholds, quality level thresholds, quality text sampler settings, and Ollama configuration).
4. `ExportReportWiring` relies on `DocxReportSettings`' default factories to fetch environment variables.

This scattered approach makes tracking, documenting, and validating configuration variables difficult, and couples individual components directly to the environment.

### Affected Areas
- `src/infrastructure/config/recommendation_config.py` — Removed. Its responsibilities are taken over by the centralized environment configuration class.
- `src/infrastructure/adapters/report/docx_report_settings.py` — Modified. Static defaults replace all `default_factory=lambda: environ.get(...)` calls to isolate the settings class from the environment.
- `src/infrastructure/wirings/analyze_document_use_case_wiring.py` — Modified. It will import the centralized configuration class, instantiate it, and use its public attributes and DTO builders to wire dependencies.
- `src/infrastructure/wirings/export_report_wiring.py` — Modified. It will import the centralized configuration class, instantiate it, and inject configuration values into `DocxReportSettings`.
- `src/infrastructure/environment_config.py` — New. Holds all configuration parsing logic, exposing public attributes and DTO constructor methods.
- `src/infrastructure/tests/test_analyze_document_use_case_wiring.py` — Modified. Updated to ensure environment settings are propagated via the new centralized class.
- `src/infrastructure/tests/adapters/report/test_docx_report_settings.py` — Modified. Replaced tests asserting dynamic environment overrides on `DocxReportSettings` with tests asserting static defaults.
- `src/infrastructure/tests/test_environment_config.py` — New. Unit tests for parsing, default values, type casting, and DTO builders in the centralized configuration class.
- `.env` & `.env.example` — Modified. Added missing keys for recommendation thresholds.

### Approaches
1. **Cached Attributes (Parsed at Instantiation)** — In this approach, `EnvironmentConfig` reads and parses all environment variables in its `__init__` constructor, converting them to the appropriate types (such as `int`, `float`, `str`) and caching them as instance attributes.
   - Pros:
     - Parses and casts environment values once, failing fast on invalid types during instantiation.
     - Strong IDE auto-completion and static analysis support via class attribute typing.
     - Follows clean architecture invariants by keeping the object state immutable and predictable once constructed.
   - Cons:
     - Changes to `os.environ` during testing will not be reflected unless a new instance is constructed (though our wiring classes are constructed per-request/run, avoiding this issue).
   - Effort: Low

2. **Dynamic `getenv` Reads (Evaluated at Call Time)** — In this approach, `EnvironmentConfig` does not parse values during construction. Instead, it exposes properties/getters that call `os.getenv` and perform type casting on every call.
   - Pros:
     - Always fetches the latest environment variable state.
     - Slightly simpler unit tests using standard `unittest.mock.patch.dict(os.environ, ...)`.
   - Cons:
     - Adds runtime parsing overhead on every call.
     - Type casting errors are deferred until runtime usage rather than fail-fast on startup.
     - Adds significant property getter boilerplate for 37 configuration fields.
   - Effort: Medium

### Recommendation
We recommend **Approach 1 (Cached Attributes)**. Centralizing config parsing at instantiation aligns with clean architecture conventions, enables type safety, and ensures that invalid environment configurations fail early (fail-fast principle). The wiring classes are instantiated per-run/use-case creation, which mitigates the risk of stale configurations during test runs.

Additionally, to strictly follow the Clean Architecture conventions of the project:
- The class MUST be named `EnvironmentConfig` and located in `src/infrastructure/environment_config.py` to avoid the abbreviation `EnvConfig` (in accordance with the project's "No abbreviations" rule).
- In accordance with the "Only stdlib specific name imports" standard, all imports from `os` must use specific names (e.g., `from os import getenv`).
- For typing, Python PEP 604 Union types (e.g., `Type | None`) must be used instead of `Optional`.

### Risks
- **Test Environment Isolation**: Standard `patch.dict(os.environ, ...)` in tests might fail if `EnvironmentConfig` instances are cached globally or reused across test methods.
  - Mitigation: Ensure `EnvironmentConfig` is instantiated within each wiring method call or use-case creation method, rather than module-level caching, or reinstantiated per test run.
- **Data Type Mismatches**: Environment variables are strings, so casting them to `int` or `float` could raise `ValueError` at runtime if unset or invalid in the environment.
  - Mitigation: All type conversions in `EnvironmentConfig` constructor must have sensible fallback default values.

### Ready for Proposal
Yes. The exploration is complete and details the design and naming decisions. We can now proceed to the design/proposal phase.
