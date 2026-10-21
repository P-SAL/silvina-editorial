# Proposal: Resolve Magic Values Debt

## Intent

Address magic values debt in domain validators, adapters, and formatting layouts. Parameterize hardcoded numbers and strings using injected configuration and environment variables.

## Scope

### In Scope
- Parameterize `StructureValidator` header length limit.
- Parameterize `DocxCitationAdapter` author name limit.
- Replace raw regex flags in `DocxReferenceAdapter`.
- Parameterize `LanguageToolAdapter` maximum replacements.
- Group report layout parameters into `DocxReportSettings`.
- Expose configurations via `.env` and `.env.example`.
- Load variables in `AnalyzeDocumentUseCaseWiring` and `ExportReportWiring`.

### Out of Scope
- Changing structural section alias keywords.
- Supporting format types other than Word (`.docx`).

## Capabilities

### New Capabilities
None

### Modified Capabilities
- `validate-structure`: Parameterize the 100-character header limit with environment variable `STRUCTURE_MAX_HEADER_LENGTH`.
- `extract-citations`: Parameterize the 100-character author name limit with environment variable `CITATION_MAX_AUTHOR_NAME_LENGTH`.
- `check-grammar`: Parameterize the 3-suggestion replacement limit with environment variable `GRAMMAR_MAX_REPLACEMENTS`.
- `export-report`: Group and parameterize formatting configurations in `DocxReportSettings` using `REPORT_WORDS_PER_PAGE`, `REPORT_MAX_ERRORS_DISPLAYED`, `REPORT_CONTEXT_TRUNCATION_LIMIT`, and `REPORT_MAX_REPLACEMENTS`.

## Approach

1. Update `StructureValidator`, `DocxCitationAdapter`, and `LanguageToolAdapter` to accept limit parameters in `__init__` with defaults.
2. In `DocxReferenceAdapter`, replace raw flags `2 | 16` with `re.IGNORECASE | re.DOTALL`.
3. In `DocxReportSettings`, add config fields and default factories to load `REPORT_` environment variables. Refactor `DocxReportAdapter` to read layout limits from settings.
4. Load new variables in wirings `AnalyzeDocumentUseCaseWiring` and `ExportReportWiring` and inject them.
5. Add configuration entries to `.env` and `.env.example`.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/domain/structure/structure_validator.py` | Modified | Add `max_header_length` parameter. |
| `src/infrastructure/adapters/document/docx_citation_adapter.py` | Modified | Add `max_author_name_length` parameter. |
| `src/infrastructure/adapters/document/docx_reference_adapter.py` | Modified | Use symbolic `re` flags. |
| `src/infrastructure/adapters/grammar/language_tool_adapter.py` | Modified | Add `max_replacements` parameter. |
| `src/infrastructure/adapters/report/docx_report_settings.py` | Modified | Expose new `REPORT_*` settings. |
| `src/infrastructure/adapters/report/docx_report_adapter.py` | Modified | Read limits from `DocxReportSettings`. |
| `src/infrastructure/wirings/analyze_document_use_case_wiring.py` | Modified | Read environment variables and inject dependencies. |
| `src/infrastructure/wirings/export_report_wiring.py` | Modified | Read environment variables and inject dependencies. |
| `.env`, `.env.example` | Modified | Define configuration defaults. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Missing environment variables crash the application | Low | Fallback to sensible defaults in code. |
| Formatting layout breakage due to custom limits | Low | Validate limits or handle truncation safely in adapters. |

## Rollback Plan

Revert git changes to the pre-change commit:
```bash
git checkout HEAD -- src/ .env.example
```

## Dependencies

- None

## Success Criteria

- [ ] All new variables are documented in `.env.example`.
- [ ] No hardcoded layout limits remain in `DocxReportAdapter`.
- [ ] Existing test suites pass successfully.
