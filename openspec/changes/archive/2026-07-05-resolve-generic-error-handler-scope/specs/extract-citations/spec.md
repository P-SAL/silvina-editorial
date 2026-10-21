# Delta Spec — extract-citations

**Change**: resolve-generic-error-handler-scope

## MODIFIED Requirements

### R7 — DocxCitationAdapter

`DocxCitationAdapter(CitationExtractionPort)` MUST exist at `src/infrastructure/adapters/document/docx_citation_adapter.py`.
Constructs `CitationDTO` (not legacy `Citation`).
Raises `CitationParsingFailed` on failure. MUST NOT import from `data_access/`.

(Previously: `@generic_error_handler` on `extract_citations` was required.)

**S7a — Valid file**: GIVEN `1. test_Científico.docx` | WHEN `extract_citations(docx_path)` called | THEN non-empty `list[CitationDTO]` returned without raising.
**S7b — Item type**: GIVEN result from S7a | WHEN each item inspected | THEN all are `CitationDTO` instances.
**S7c — Citation type**: GIVEN result from S7a | WHEN `citation_type` accessed on any item | THEN `CitationType.AUTHOR_YEAR`.

### R8 — DocxReferenceAdapter

`DocxReferenceAdapter(ReferenceExtractionPort)` MUST exist at `src/infrastructure/adapters/document/docx_reference_adapter.py`.
`section_type` MUST be one of `{"Bibliografía", "Referencias", "Fuentes bibliográficas"}`.
Raises `ReferenceParsingFailed` on failure. MUST NOT import from `data_access/`.

(Previously: `@generic_error_handler` on `extract_references` was required.)

**S8a — Valid file**: GIVEN `1. test_Científico.docx` | WHEN `extract_references(docx_path)` called | THEN `(non-empty list[ReferenceDTO], non-empty str)` returned without raising.
**S8b — Item type**: GIVEN returned list | WHEN items inspected | THEN all are `ReferenceDTO` instances.
**S8c — section_type**: GIVEN returned string | THEN value is one of the three allowed section names.
