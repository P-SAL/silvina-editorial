# Proposal: docx-citation-location-fix

## Intent

Resolve the bug where `DocxCitationAdapter` returns `location=-1` for all extracted citations. This causes APA citation validation reports to show empty paragraph previews in production.

## Scope

### In Scope
- Modify `DocxCitationAdapter` to scan documents paragraph by paragraph, tracking the 0-based index.
- Assign the correct paragraph index to the `location` field in each extracted `CitationDTO`.
- Support fallback to raw text matching with `location=0` for backward compatibility.
- Test changes through the public interface (`extract_citations`).

### Out of Scope
- Modifying `ApaValidator` behavior or output.
- Writing unit tests targeting private methods.
- Changing citation deduplication logic (document-wide `seen` set remains).

## Capabilities

### New Capabilities
None

### Modified Capabilities
- `extract-citations`: `DocxCitationAdapter` must calculate and assign the correct 0-based paragraph index to the `location` field of each extracted `CitationDTO`.

## Approach

1. Update `DocxCitationAdapter._extract_citations(self, paragraphs: list[str] | str, full_text: str | None = None)` signature to support both standard paragraph list and raw text fallback:
   ```python
   if full_text is not None:
       paragraphs = [full_text]
   elif isinstance(paragraphs, str):
       paragraphs = [paragraphs]
   ```
2. Loop over paragraphs with `enumerate(paragraphs)`.
3. In helper methods (`_collect_parenthetical`, `_collect_multi_author`, `_collect_single_author`), scan each paragraph individually and set `location=p_idx`.
4. Keep the document-wide `seen` set to preserve current deduplication.
5. In `extract_citations`, call `_extract_citations(paragraphs)`.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/infrastructure/adapters/document/docx_citation_adapter.py` | Modified | Update signature and loop logic to pass paragraph index to DTOs. |
| `src/infrastructure/tests/adapters/document/test_docx_citation_adapter.py` | Modified | Add unit tests using public `extract_citations` validating correct `location` field. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Regex performance on many paragraphs | Low | Regex patterns are precompiled; typical document size is small. |
| Legacy test breaks due to signature change | Low | Fallback handles `full_text` gracefully, mapping to a single paragraph. |

## Rollback Plan

Revert all changes to `src/infrastructure/adapters/document/docx_citation_adapter.py` and `src/infrastructure/tests/adapters/document/test_docx_citation_adapter.py` using Git.

## Dependencies

- None

## Success Criteria

- [ ] Citation validation reports show non-empty paragraph previews.
- [ ] `DocxCitationAdapter.extract_citations` maps citations to their correct 0-based paragraph index.
- [ ] All existing and new tests pass successfully.
