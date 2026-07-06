## Exploration: resolve-magic-values-debt

### Current State

The codebase contains several hardcoded magic numbers, magic strings, and raw regex flags across its domain validators and infrastructure adapters:

1. **`StructureValidator`** (`src/domain/structure/structure_validator.py`):
   - In `_extract_present_sections()`, the character limit for a short header line is hardcoded as `100` (`is_short_header = len(text_lower) < 100`).

2. **`DocxCitationAdapter`** (`src/infrastructure/adapters/document/docx_citation_adapter.py`):
   - In `_collect_multi_author()`, the maximum length of an author name is hardcoded as `100` (`if len(author) > 100 or author.startswith(_INTRO_PHRASES):`).

3. **`DocxReferenceAdapter`** (`src/infrastructure/adapters/document/docx_reference_adapter.py`):
   - Raw regular expression flags `2 | 16` are used when compiling `_BIB_SECTION_PATTERN` instead of symbolic names `re.IGNORECASE | re.DOTALL`.

4. **`LanguageToolAdapter`** (`src/infrastructure/adapters/grammar/language_tool_adapter.py`):
   - In `_map_to_dto()`, the number of grammar suggestions returned is hardcoded to `3` (`replacements=match.replacements[:3],`).

5. **`DocxReportAdapter`** (`src/infrastructure/adapters/report/docx_report_adapter.py`):
   - Several magic formatting/report values are hardcoded:
     - Divisor `250` for estimating pages (`estimated_pages = doc_content.word_count // 250`).
     - List slice size `5` for limiting the number of displayed grammar errors (`grammar.errors[:5]`) and APA validation violations (`errors[:5]`).
     - Context truncation limit of `150` characters (`context_text = err.context if len(err.context) < 150 else err.context[:150] + "..."`).
     - Suggestion replacements limit of `3` (`err.replacements[:3]`).
     - Metrics table dimensions of `6` rows by `2` columns (`table = doc.add_table(rows=6, cols=2)`).
     - Footer separator line character repetition count of `80` (`"─" * 80`).

### Affected Areas

- [structure_validator.py](file:///E:/Python/silvina-editorial/src/domain/structure/structure_validator.py#L49) — parameterize hardcoded 100-character header limit.
- [docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_citation_adapter.py#L99) — parameterize hardcoded 100-character maximum author name limit.
- [docx_reference_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_reference_adapter.py#L22) — replace raw regex flags `2 | 16` with `IGNORECASE | DOTALL`.
- [language_tool_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/grammar/language_tool_adapter.py#L55) — parameterize hardcoded 3-replacement limit.
- [docx_report_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/report/docx_report_adapter.py#L228) — use config fields from `DocxReportSettings` instead of hardcoded report-specific layout/dimension values.
- [docx_report_settings.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/report/docx_report_settings.py) — group the report configurations/layout constraints into the existing dataclass, loading them via environment variables.
- [analyze_document_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py) — load the environment variables and pass them into the constructors of `StructureValidator`, `DocxCitationAdapter`, and `LanguageToolAdapter`.
- [export_report_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/export_report_wiring.py) — load environment variables to supply `DocxReportSettings`.
- `.env.example` / `.env` — define the new environment variables with their defaults.

### Approaches

1. **Parameter & DTO Configuration (Recommended)**
   - Add parameters with default values to constructors for adapters/validators with single magic values (`StructureValidator`, `DocxCitationAdapter`, `LanguageToolAdapter`).
   - Group the report-specific layout variables (words/page, list limits, context size, etc.) into the existing `DocxReportSettings` class.
   - Load all environment variables in the wirings (`AnalyzeDocumentUseCaseWiring` and `ExportReportWiring`) and pass them.
   - Update `DocxReferenceAdapter` to import `IGNORECASE` and `DOTALL` from `re` and use `flags=IGNORECASE | DOTALL`.
   - **Pros**: Fits the existing design pattern (e.g. `ArticleSizeThresholdsDTO`), ensures backwards compatibility (due to defaults), and keeps instantiation logic in the composition root (wiring).
   - **Cons**: Requires minor updates in multiple wirings.
   - **Effort**: Low

2. **Dedicated Config DTO Class per Class**
   - Create separate config DTO files and classes for every validator and adapter (e.g., `StructureValidatorConfigDTO`, `LanguageToolConfigDTO`, etc.) and inject them.
   - **Pros**: Uniform config patterns for all classes.
   - **Cons**: Overkill for classes that only have one configurable parameter (e.g. `max_header_length`), creating unnecessary boilerplate files.
   - **Effort**: Medium

### Recommendation

We recommend **Approach 1**. Injecting single constructor parameters (with sensible defaults) is Pythonic and simple for classes with a single configuration option. For the report adapter, grouping the various formatting parameters into the existing `DocxReportSettings` object is a natural extension of its purpose.

### Risks

- **Backwards Compatibility**: Ensure that default constructor parameters are provided so that tests and legacy instantiations continue to function without changes.
- **Environment Variables**: Document and provide clear defaults for all new environment variables in `.env` and `.env.example`.

### Ready for Proposal
Yes — the next step is to create the formal design proposal and task list in `openspec/changes/resolve-magic-values-debt/`.
