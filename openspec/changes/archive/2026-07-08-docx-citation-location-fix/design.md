# Design: docx-citation-location-fix

## Technical Approach

To resolve the issue where validation reports show empty paragraph previews (caused by `location=-1` for all extracted citations), we will update `DocxCitationAdapter` to scan documents paragraph by paragraph. Instead of joining paragraphs into a single text block, we will:
1. Retain the paragraph list returned by the document text adapter.
2. Iterate over paragraphs sequentially, tracking the 0-based paragraph index (`p_idx`).
3. Set the `location` field of `CitationDTO` to the paragraph's index `p_idx` during iteration.
4. Support legacy string/full-text inputs (for backward compatibility in tests) by converting them to a single-element list `[full_text]` and mapping to location `0`.

This approach ensures correct paragraph preview generation while preserving the current extraction, filtering, and deduplication logic.

## Architecture Decisions

### Decision: Paragraph Index Tracking Strategy

| Option | Tradeoff | Decision |
|---|---|---|
| Character Index Mapping | High complexity; requires mapping match offsets to paragraph boundaries. Risk of off-by-one errors. | Rejected |
| Paragraph-by-Paragraph Scan | Loops regex matches over paragraphs. Simple, clean, directly maps loop index to `location`. Highly performant for normal document sizes. | **Accepted** |

### Decision: Compatibility and Fallback Interface

| Option | Tradeoff | Decision |
|---|---|---|
| Break private signature of `_extract_citations` | Breaks existing tests that call the helper directly with text strings. | Rejected |
| Support both lists and strings via fallback | Minimal signature change, automatically maps legacy string/full_text invocations to a single paragraph list `[full_text]`, assigning location `0`. | **Accepted** |

## Data Flow

Data flow transitions from a single concatenated string to sequential paragraph scanning:

```mermaid
graph TD
    A[DocxTextAdapter] -->|read_paragraphs| B[DocxCitationAdapter.extract_citations]
    B -->|pass list[str]| C[DocxCitationAdapter._extract_citations]
    C -->|enumerate paragraphs| D[Private Helper Methods]
    D -->|match citation in p_idx| E[Create CitationDTO]
    E -->|location = p_idx| F[Result List]
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/infrastructure/adapters/document/docx_citation_adapter.py` | Modify | Update constructor to make `max_author_name_length` optional (default `100`). Modify `extract_citations` to pass paragraphs list. Update `_extract_citations` and internal collection helpers to loop over paragraphs and set `location=p_idx`. |
| `src/infrastructure/tests/adapters/document/test_docx_citation_adapter.py` | Modify | Add tests verifying that `location` corresponds to the correct 0-based paragraph index (Scenario S7e) and fallback maps to `0` (Scenario S7f). |

## Interfaces / Contracts

No public interface changes. The private signature of `_extract_citations` is updated:

```python
class DocxCitationAdapter(CitationExtractionPort):
    def __init__(
        self,
        document_text_port: DocumentTextPort,
        max_author_name_length: int = 100,
    ) -> None:
        ...

    def _extract_citations(
        self,
        paragraphs: list[str] | str | None = None,
        full_text: str | None = None,
    ) -> list[CitationDTO]:
        ...
```

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | Paragraph index tracking | Assert `CitationDTO.location` equals the exact paragraph index in mock documents (Scenario S7e). |
| Unit | Fallback behavior | Pass raw text/full_text and assert `CitationDTO.location` is 0 (Scenario S7f). |
| Integration | E2E citation validation | Execute `test_docx_citation_adapter.py` suite against sample docx documents to ensure no regressions. |

## Migration / Rollout

No migration required.
