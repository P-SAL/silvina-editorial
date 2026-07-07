# Design: Centralize Configuration

## Technical Approach

We will centralize all environment variable configurations into a single, concrete `EnvConfig` class in `src/infrastructure/env_config.py`. At instantiation, this class reads, casts, and caches all environment variables as public attributes. We will:
1. Update `AnalyzeDocumentUseCaseWiring` and `ExportReportWiring` to instantiate `EnvConfig` and inject settings.
2. Refactor `DocxReportSettings` to have static defaults and accept constructor injections.
3. Delete `recommendation_config.py`.
4. Update/create tests in `test_env_config.py`, `test_analyze_document_use_case_wiring.py`, and `test_docx_report_settings.py`.

## Architecture Decisions

| Option | Tradeoff | Decision |
|---|---|---|
| Instantiation-time parsing (Cached Attributes) | Throws type errors early, clear typing for attributes, but static values won't track subsequent changes to `os.environ` unless re-instantiated. | **Chosen**. Aligns with clean architecture, failing fast on bad config. |
| Call-time parsing (Dynamic `getenv` reads) | Tracks runtime environment modifications dynamically, but adds overhead, delays failures, and requires boilerplate. | Rejected. |

| Option | Tradeoff | Decision |
|---|---|---|
| Injecting settings into `DocxReportSettings` | Clear dependency injection from the wiring layer, but requires modifying the signature of `DocxReportSettings`. | **Chosen**. Keeps the adapter class clean and independent of `os.environ`. |
| Keeping `default_factory` inside `DocxReportSettings` | Minimal edits, but couples the adapter to `os` environment directly. | Rejected. |

## Data Flow
The wiring class instantiates `EnvConfig` and propagates configuration values down the dependency tree.

```
       [Environment]
             │
             ▼
        [EnvConfig]
             │
      ┌──────┴────────────────────────┐
      ▼                               ▼
[AnalyzeDocumentUseCaseWiring]   [ExportReportWiring]
      │                               │
      ├─► (Parses DTOs / attributes)  ├─► [DocxReportSettings] (Injected)
      ▼                               ▼
[AnalyzeDocumentUseCase]         [DocxReportAdapter]
                                      │
                                      ▼
                               [ExportReportUseCase]
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/infrastructure/env_config.py` | Create | Contains the `EnvConfig` class parsing environment variables and building the `RecommendationSettingsDTO`. |
| `src/infrastructure/tests/test_env_config.py` | Create | Unit tests for environment variable loading, defaults, type-casting, and recommendation settings creation. |
| `src/infrastructure/config/recommendation_config.py` | Delete | Replaced by `EnvConfig`. |
| `src/infrastructure/wirings/analyze_document_use_case_wiring.py` | Modify | Instantiates `EnvConfig` and passes config values to private wiring methods. |
| `src/infrastructure/wirings/export_report_wiring.py` | Modify | Instantiates `EnvConfig` and injects settings into the `DocxReportSettings` constructor, checking for `python-docx` availability. |
| `src/infrastructure/adapters/report/docx_report_settings.py` | Modify | Replaces `default_factory` environment calls with static defaults, isolating it from `os.environ`. |
| `src/infrastructure/tests/adapters/report/test_docx_report_settings.py` | Modify | Updates tests to verify static defaults and construction overrides without environment mocking. |
| `src/infrastructure/tests/test_analyze_document_use_case_wiring.py` | Modify | Updates tests to ensure the new centralized config is correctly wired and injected. |

## Interfaces / Contracts

```python
# src/infrastructure/env_config.py
from os import getenv
from src.domain.dtos.recommendation_settings_dto import RecommendationSettingsDTO

class EnvConfig:
    def __init__(self) -> None:
        # Loaded public attributes:
        self.citation_max_author_name_length: int = int(getenv("CITATION_MAX_AUTHOR_NAME_LENGTH", "100"))
        # (Other 36 parsed environment variables mapped similarly)
        ...

    def get_recommendation_settings(self) -> RecommendationSettingsDTO: ...
```

```python
# src/infrastructure/adapters/report/docx_report_settings.py
from dataclasses import dataclass

@dataclass(frozen=True)
class DocxReportSettings:
    app_name: str = "Silvina Editorial Assistant"
    app_version: str = "0.9"
    score_high_threshold: float = 8.0
    score_medium_threshold: float = 6.0
    words_per_page: int = 250
    max_errors_displayed: int = 5
    context_truncation_limit: int = 150
    max_replacements: int = 3
    # Visual design settings omitted for brevity
```

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | `EnvConfig` | Assert default fallback values when env is empty, verify casting, and validate DTO output. |
| Unit | `DocxReportSettings` | Assert static defaults without patched environment, and verify constructor argument overriding. |
| Integration | `AnalyzeDocumentUseCaseWiring` | Verify all dependencies are wired with correct config. |
| Integration | `ExportReportWiring` | Verify wiring uses `EnvConfig` overrides, and throws `ReportExportUnavailable` when `python-docx` is missing. |

## Migration / Rollout
No migration required.

## Open Questions
- None.
