## Exploration: Fix paragraph_preview always empty by resolving the location=-1 bug in DocxCitationAdapter so that it calculates the correct paragraph index.

### Current State
In [docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_citation_adapter.py), when citations are extracted, the `extract_citations` method fetches a list of paragraphs from the document text port, joins them with a space to form a single string `full_text`, and runs regex searches on this `full_text`.
Because it processes a single joined string, the link to the original paragraph indices is lost. As a result, the adapter instantiates all `CitationDTO` objects with `location=-1`.
Consequently, when [ApaValidator.validate_all_citations](file:///E:/Python/silvina-editorial/src/domain/citation/apa_validator.py#L32-L46) iterates over the extracted citations, it performs a bounds check on `citation.location`:
`paragraphs[citation.location] if 0 <= citation.location < len(paragraphs) else ""`
Since `location` is always `-1`, the check fails, `paragraph_text` defaults to `""`, and `paragraph_preview` is returned as an empty string for all `AUTHOR_YEAR` citations in production reports.

### Affected Areas
- [docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_citation_adapter.py) — Needs modification to parse text paragraph by paragraph and track the 0-based paragraph index (`location`).
- [test_docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/adapters/document/test_docx_citation_adapter.py) — Needs a new unit test verifying that correct paragraph indices are calculated and populated on `CitationDTO`.

### Approaches
1. **Paragraph-by-Paragraph Regex Scan (Recommended)**
   Scan each paragraph in the document independently by passing the paragraph list to the private helper methods `_collect_parenthetical`, `_collect_multi_author`, and `_collect_single_author`. The 0-based index `p_idx` is then naturally mapped to the `location` parameter of each `CitationDTO`. Backward compatibility is preserved by accepting a fallback `paragraphs=None` which defaults to `[full_text]` (e.g. in tests where `_extract_citations` is called directly on raw text).
   - **Pros:**
     - Directly resolves the `location=-1` bug.
     - Naturally tracks the paragraph index using the loop counter.
     - Avoids potential false-positive cross-paragraph matches when joining paragraphs.
     - Highly localized changes in `DocxCitationAdapter` with no impact on other files.
   - **Cons:**
     - Runs the compiled regex patterns multiple times (once per paragraph) instead of once per document, though this has negligible performance impact.
   - **Effort:** Low

2. **Full Text Match Index Translation**
   Perform the regex scan on the joined `full_text` once, but use `finditer()` to obtain the character indices of each match. Maintain a mapping of character offsets back to paragraph indices, and translate each match's character start index to a paragraph index using a binary search.
   - **Pros:**
     - Executes regex matching only once per document.
   - **Cons:**
     - Highly complex offset translation logic.
     - Prone to bugs and edge cases (e.g. when citations span across paragraphs due to joining spacing).
   - **Effort:** Medium

### Recommendation
Approach 1 is recommended. It is simpler, much easier to test, more robust against cross-paragraph false positives, and preserves perfect backward compatibility with existing unit tests that pass a single string to `_extract_citations`.

### Risks
- **Regex performance:** Iterating through paragraphs one by one is fast enough for typical academic papers but could be slower if a document has thousands of paragraphs. However, since the regex patterns are compiled once and python-docx text loading is fast, the impact is negligible.
- **Unit Test Compatibility:** Unit tests directly call the private `_extract_citations(full_text=...)` method. We must ensure that when `paragraphs` is omitted, the method gracefully defaults to `[full_text]`, assigning `location=0` so that existing assertions continue to pass without changes.

### Ready for Proposal
Yes. The orchestrator is ready to create the proposal for this change.
