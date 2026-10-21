# Delta Spec — extract-citations

**Change**: resolve-magic-values-debt

## MODIFIED Requirements

### R7 — DocxCitationAdapter

`DocxCitationAdapter` MUST accept an optional `max_author_name_length: int` in its constructor (defaulting to 100) and store it. During multi-author extraction, if the author name length exceeds `max_author_name_length`, it MUST NOT be parsed as an author.
`DocxCitationAdapter(CitationExtractionPort)` MUST exist at `src/infrastructure/adapters/document/docx_citation_adapter.py`.
Constructs `CitationDTO` (not legacy `Citation`).
Raises `CitationParsingFailed` on failure. MUST NOT import from `data_access/`.

(Previously: The maximum author name length limit of 100 was hardcoded.)

**S7a — Valid file**: GIVEN `1. test_Científico.docx` | WHEN `extract_citations(docx_path)` called | THEN non-empty `list[CitationDTO]` returned without raising.
**S7b — Item type**: GIVEN result from S7a | WHEN each item inspected | THEN all are `CitationDTO` instances.
**S7c — Citation type**: GIVEN result from S7a | WHEN `citation_type` accessed on any item | THEN `CitationType.AUTHOR_YEAR`.
**S7d — Custom max_author_name_length**: GIVEN a `DocxCitationAdapter` initialized with `max_author_name_length` = 5 | WHEN a citation with a long author name is processed | THEN it is rejected and not returned.

---

### R8 — DocxReferenceAdapter

`DocxReferenceAdapter(ReferenceExtractionPort)` MUST exist at `src/infrastructure/adapters/document/docx_reference_adapter.py`.
`section_type` MUST be one of `{"Bibliografía", "Referencias", "Fuentes bibliográficas"}`.
Raises `ReferenceParsingFailed` on failure. MUST NOT import from `data_access/`.
Regex patterns MUST use symbolic flags (e.g. `re.IGNORECASE | re.DOTALL`) instead of raw integers (e.g. `2 | 16`).

(Previously: Raw integer values were used for regex compile flags.)

**S8a — Valid file**: GIVEN `1. test_Científico.docx` | WHEN `extract_references(docx_path)` called | THEN `(non-empty list[ReferenceDTO], non-empty str)` returned without raising.
**S8b — Item type**: GIVEN returned list | WHEN items inspected | THEN all are `ReferenceDTO` instances.
**S8c — section_type**: GIVEN returned string | THEN value is one of the three allowed section names.

---

### R10 — AnalyzeDocumentUseCaseWiring Wires the Adapters Directly

`AnalyzeDocumentUseCaseWiring._get_citation_extraction_port()` MUST return a `DocxCitationAdapter` configured with `max_author_name_length` read from the environment variable `CITATION_MAX_AUTHOR_NAME_LENGTH` (defaulting to 100).
`_get_reference_extraction_port()` MUST return a `DocxReferenceAdapter` (both constructed with the shared `_get_document_text_port()` instance).

(Previously: `_get_citation_extraction_port()` initialized `DocxCitationAdapter` with only `document_text_port` and no author length limit.)

**S10a — Correct types**: GIVEN `AnalyzeDocumentUseCaseWiring().create_use_case()` called | WHEN
the resulting use case's private attributes are inspected | THEN
`isinstance(uc._citation_extraction_port, DocxCitationAdapter)` and
`isinstance(uc._reference_extraction_port, DocxReferenceAdapter)` are both `True`.
