# Design: resolve-magic-values-debt

## Technical Approach

This design resolves magic values debt by parameterizing hardcoded thresholds, regex flags, and report layout constraints. By moving all magic numbers and strings to parameters resolved in the dependency injection wirings (Composition Root), the application's core domain services and infrastructure adapters become fully configurable.

## Architecture Decisions

| Area | Option | Tradeoff | Decision |
|---|---|---|---|
| **Parameterization Pattern** | Constructor parameters vs. config DTO classes | Constructor params are simpler and prevent boilerplate for classes with only one config value. DTOs scale better for large groups. | Use constructor parameters for single-config classes (`StructureValidator`, `DocxCitationAdapter`, `LanguageToolAdapter`). Use `DocxReportSettings` to group report layout parameters. |
| **Configuration Loading** | Access `os.getenv` directly in adapters vs. Injecting in Wiring | Direct access adds environment side-effects and makes unit testing harder. Wiring loading keeps the classes pure and testable. | Load configurations in `AnalyzeDocumentUseCaseWiring` and `ExportReportWiring`, then inject them. |
| **Regex Flags** | Raw integer flags (`2 \| 16`) vs. Symbolic flags | Integer flags are opaque and non-standard. Symbolic flags are self-documenting. | Import and use `IGNORECASE \| DOTALL` from the `re` module. |
| **Docx Adapter Availability** | Soft-check at runtime vs. Guard on initialization | Importing `docx` at module level will cause runtime `ImportError` on startup. Try-except wrapping lets us handle it gracefully. | Use try-except import guard setting a module-level `DOCX_AVAILABLE` boolean. Raise `ReportExportUnavailable` on `DocxReportAdapter.__init__`. |

## Data Flow

Environment variables are loaded from `.env` on application startup. The wiring classes read these values and pass them as constructor parameters to the respective instances.

```
[.env] ──> [load_dotenv()] ──> [Wiring (Composition Root)]
                                         │
                                         ▼ (instantiate & inject)
                        [Validators / Adapters (with defaults)]
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/domain/structure/structure_validator.py` | Modify | Add `max_header_length: int = 100` to constructor. Use parameter in header length check. |
| `src/infrastructure/adapters/document/docx_citation_adapter.py` | Modify | Add `max_author_name_length: int = 100` to constructor. Use parameter in multi-author validation check. |
| `src/infrastructure/adapters/document/docx_reference_adapter.py` | Modify | Import `IGNORECASE` and `DOTALL` from `re`. Replace flags `2 \| 16` with `IGNORECASE \| DOTALL`. |
| `src/infrastructure/adapters/grammar/language_tool_adapter.py` | Modify | Add `max_replacements: int` to constructor (no default; sole default lives in the wiring's env read). Keep the existing manual try/except raising `GrammarCheckUnavailable` in `check()`. Slice replacements in `GrammarErrorDTO`. |
| `src/infrastructure/adapters/report/docx_report_settings.py` | Modify | Add `words_per_page`, `max_errors_displayed`, `context_truncation_limit`, and `max_replacements` as fields with default factories. |
| `src/infrastructure/adapters/report/docx_report_adapter.py` | Modify | Conditionally import `docx`. Raise `ReportExportUnavailable` in `__init__` if `DOCX_AVAILABLE` is `False`. Accept `settings` parameter. Update rendering logic to use settings. |
| `src/infrastructure/wirings/analyze_document_use_case_wiring.py` | Modify | Read environment variables `STRUCTURE_MAX_HEADER_LENGTH`, `CITATION_MAX_AUTHOR_NAME_LENGTH`, and `GRAMMAR_MAX_REPLACEMENTS`, and pass them. |
| `src/infrastructure/wirings/export_report_wiring.py` | Modify | Instantiate `DocxReportSettings` and pass to `DocxReportAdapter`. No `load_dotenv` call needed here — `AnalyzeDocumentUseCaseWiring`'s module-level `load_dotenv()` already runs first in every real entry point (`main.py`, `gradio_app.py`). |
| `.env` | Modify | Define environment variables for the new configurations. |
| `.env.example` | Modify | Document environment variables for the new configurations. |

## Interfaces / Contracts

```python
# src/domain/structure/structure_validator.py
class StructureValidator:
    def __init__(self, max_header_length: int) -> None:
        self._max_header_length = max_header_length

# src/infrastructure/adapters/document/docx_citation_adapter.py
class DocxCitationAdapter(CitationExtractionPort):
    def __init__(
        self,
        document_text_port: DocumentTextPort,
        max_author_name_length: int,
    ) -> None:
        self._document_text_port = document_text_port
        self._max_author_name_length = max_author_name_length

# src/infrastructure/adapters/grammar/language_tool_adapter.py
from language_tool_python import LanguageTool

class LanguageToolAdapter(GrammarCheckPort):
    def __init__(self, max_replacements: int, language: str = "es") -> None:
        self._language = language
        self._max_replacements = max_replacements
        self._tool: LanguageTool | None = None

# src/infrastructure/adapters/report/docx_report_adapter.py
from src.domain.exceptions.report_errors import ReportExportUnavailable

class DocxReportAdapter(ReportExportPort):
    def __init__(
        self,
        logo_path: str | None = None,
        settings: DocxReportSettings | None = None,
    ) -> None:
        if not DOCX_AVAILABLE:
            raise ReportExportUnavailable()
        self._logo_path = logo_path
        self._settings = settings or DocxReportSettings()
```

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | Custom limit filters | Test `StructureValidator`, `DocxCitationAdapter`, and `LanguageToolAdapter` with custom limits to verify correct filtering and truncation. |
| Integration | Config mapping | Test that wirings fetch from environment and construct adapters correctly. |
| Regression | Export settings | Test `DocxReportAdapter` layout settings (e.g. limiting errors, page estimation) are respected. |
| Regression | Availability Check | Verify that mock `DOCX_AVAILABLE = False` raises `ReportExportUnavailable` on `DocxReportAdapter` construction. |

## Migration / Rollout

No database or schema migration required. Ensure production settings contain default or custom values for the environment variables.

## Open Questions

None.
