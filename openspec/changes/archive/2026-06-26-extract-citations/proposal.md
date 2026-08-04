# Proposal: extract-citations (Slice 7 — Hexagonal Migration)

**Change name**: extract-citations
**Slice**: 7 of N (incremental hexagonal migration)
**Date**: 2026-06-26
**Status**: proposed

---

## 1. Intent

### Problem

Two legacy classes — `data_access/citation_parser.py` (`CitationParser`) and
`data_access/reference_parser.py` (`ReferenceParser`) — perform DOCX I/O directly
inside the legacy layer. They mix file-reading concerns (zipfile, XML parsing) with
business-logic data shaping (regex, DTO construction) and return legacy objects
(`Citation`, `Reference`) that the hexagonal domain cannot consume without coupling
to the legacy namespace.

The hexagonal `src/domain/citation/citation_matcher.py` already expects
`list[CitationDTO]` and `list[ReferenceDTO]` as inputs. There is currently no
hexagonal path to produce those lists from a `.docx` file. Any caller that needs
citation extraction today must reach across the architecture boundary into
`data_access/` and then manually bridge the legacy types, which violates the
dependency rule and prevents the system from being tested or swapped at the
infrastructure boundary.

### Why now

Slice 7 is the natural successor to Slices 5 and 6 (document text and content
extraction). Those slices established the port/adapter pattern for document I/O.
Citation and reference extraction are the remaining document-I/O concerns in the
migration plan. Delivering this slice closes the gap between the hexagonal
citation-matching domain logic and the file system, making the citation-validation
feature fully operable through the hexagonal path.

### Success criteria

- A caller in `src/` can invoke `ExtractCitationsUseCase.execute(docx_path)` and
  receive a `CitationExtractionResultDTO` with typed citations, references, and
  section type — without importing anything from `data_access/`.
- All new code is covered by `unittest.TestCase` tests using the project's
  established fake-port and integration-test patterns.
- The legacy `data_access/` classes remain untouched and the system continues to
  work end-to-end (coexistence invariant).

---

## 2. Scope

### Files to create (19 new files, 0 modified)

**Domain — ports**
- `src/domain/document/citation_extraction_port.py`
  `CitationExtractionPort(ABC)` with `extract_citations(docx_path: str) -> list[CitationDTO]`
- `src/domain/document/reference_extraction_port.py`
  `ReferenceExtractionPort(ABC)` with `extract_references(docx_path: str) -> tuple[list[ReferenceDTO], str]`

**Domain — exceptions**
- `src/domain/exceptions/reference_errors.py`
  `ReferenceError(BaseSrcError)` + `ReferenceParsingFailed(ReferenceError)`

**Domain — DTOs**
- `src/domain/dtos/citation_extraction_result_dto.py`
  `CitationExtractionResultDTO(BaseDTO)` — frozen dataclass with
  `citations: list[CitationDTO]`, `references: list[ReferenceDTO]`, `section_type: str`

**Infrastructure — adapters**
- `src/infrastructure/adapters/document/docx_citation_adapter.py`
  `DocxCitationAdapter(CitationExtractionPort)` — ports inline the XML extraction
  logic from `CitationParser.extract_from_docx()`; raises `DocumentNotFound` or
  `CitationParsingFailed`; `@generic_error_handler` on `extract_citations()`
- `src/infrastructure/adapters/document/docx_reference_adapter.py`
  `DocxReferenceAdapter(ReferenceExtractionPort)` — ports inline the regex-on-XML
  logic from `ReferenceParser.parse_from_docx()`; raises `DocumentNotFound` or
  `ReferenceParsingFailed`; `@generic_error_handler` on `extract_references()`

**Application — use case**
- `src/application/extract_citations_use_case.py`
  `ExtractCitationsUseCase` with `@generic_error_handler` on
  `execute(docx_path: str) -> CitationExtractionResultDTO`; injected ports are
  `_citation_port: CitationExtractionPort` and `_reference_port: ReferenceExtractionPort`

**Infrastructure — wiring**
- `src/infrastructure/wirings/extract_citations_use_case_wiring.py`
  `ExtractCitationsUseCaseWiring.create_use_case() -> ExtractCitationsUseCase`

**Tests — domain**
- `src/domain/tests/document/test_citation_extraction_port.py`
- `src/domain/tests/document/test_reference_extraction_port.py`
- `src/domain/tests/document/fake_citation_extraction_port.py`
- `src/domain/tests/document/fake_reference_extraction_port.py`
- `src/domain/tests/exceptions/test_reference_error.py`
- `src/domain/tests/exceptions/test_reference_parsing_failed.py`
- `src/domain/tests/dtos/test_citation_extraction_result.py`

**Tests — infrastructure**
- `src/infrastructure/tests/adapters/document/test_docx_citation_adapter.py`
- `src/infrastructure/tests/adapters/document/test_docx_reference_adapter.py`
- `src/application/tests/test_extract_citations_use_case.py`
- `src/infrastructure/tests/test_extract_citations_use_case_wiring.py`

---

## 3. Out of scope

- `extract_footnotes()` — zero uses in `src/`; called only from the legacy
  `business_logic/CitationMatcher.extract_all_citations()`, which is not in the
  hexagonal path and is not being migrated in this slice.
- Structured reference parsing — `ReferenceDTO.text` is the only field populated;
  sub-fields (`authors`, `year`, `title`, `source`) remain `None`. The
  `CitationMatcher` domain logic parses from `reference.text` itself.
- `SectionName` enum coercion — the adapter returns raw strings
  (`"Bibliografía"`, `"Fuentes bibliográficas"`, `"Referencias"`). Mapping to a
  typed enum is the caller's responsibility and is deferred to a future slice if needed.
- Any modification to `data_access/citation_parser.py` or
  `data_access/reference_parser.py` — they remain unchanged.
- Any modification to `src/domain/citation/citation_matcher.py` or its callers —
  wiring the use case into the application entry points is a future integration step.
- `__init__.py` changes — follow project convention: only add if the convention
  requires it (verify at spec time).

---

## 4. Approach

### Architecture pattern

Follows the hexagonal (ports-and-adapters) pattern established in Slices 5 and 6:

```
[docx file]
    |
    v
DocxCitationAdapter  ──implements──>  CitationExtractionPort (ABC, domain)
DocxReferenceAdapter ──implements──>  ReferenceExtractionPort (ABC, domain)
    |
    v
ExtractCitationsUseCase.execute(docx_path)
    └── calls _citation_port.extract_citations(docx_path)
    └── calls _reference_port.extract_references(docx_path)
    └── returns CitationExtractionResultDTO(citations, references, section_type)
    |
    v
ExtractCitationsUseCaseWiring  ──wires──>  DocxCitationAdapter + DocxReferenceAdapter
```

**Port placement**: `src/domain/document/` — citations are a document I/O concern,
not an LLM concern (consistent with existing ports in this package).

**Adapter logic**: ported inline from the legacy classes using `xml.etree.ElementTree`
(citations) and raw regex-on-XML (references). No imports from `data_access/`.

**Error handling**: `@generic_error_handler` applied to adapter public methods and
`use_case.execute()` only. ABC methods carry no decorator.

**DTO construction**: adapters are responsible for constructing `CitationDTO` /
`ReferenceDTO` from raw strings extracted from the XML — no legacy type bridging.

### TDD order (strict TDD mode active)

For each new class, the sequence is:
1. Write a failing test (RED)
2. Write the minimum implementation to pass (GREEN)
3. Refactor if needed (REFACTOR)

Recommended order within the slice:
1. Exception classes + tests (no dependencies)
2. Result DTO + test
3. Port ABCs + tests + fake doubles
4. Adapter integration tests (RED against real sample docx) → adapter implementations
5. Use case unit tests (with fakes) → use case implementation
6. Wiring test → wiring implementation

---

## 5. Risks

| Risk | Likelihood | Mitigation |
|------|-----------|------------|
| XML regex logic behaves differently from legacy under edge-case DOCX structures | Medium | Integration tests against `docs/sample-documents/1. test_Científico.docx` catch divergence before merge |
| Sample document has no citations or no references section | Low | Verify at spec time; if missing, a second fixture may be needed |
| `ReferenceDTO` partial population (`text`-only) causes downstream failures | Low | `CitationMatcher` already parses from `reference.text`; no callers rely on structured sub-fields in the hexagonal path today |
| Adapter XML namespaces differ between DOCX versions | Low | Port exactly from legacy; do not redesign the parsing approach |
| `@generic_error_handler` masking real errors during adapter development | Low | Run integration tests with real docx before relying on the decorator |

---

## 6. Definition of Done

Per `docs/plan-migracion-hexagonal.md` §8 — a slice is done when all of the
following are checked:

1. Entities/DTOs/enums in `src/domain/` covered by `unittest.TestCase`.
2. Port defined in `src/domain/document/` + adapter in
   `src/infrastructure/adapters/document/` with `@generic_error_handler`.
3. Use case in `src/application/` depending only on `src/domain/`.
4. Wiring in `src/infrastructure/wirings/extract_citations_use_case_wiring.py`.
5. Wiring test in `src/infrastructure/tests/` asserting correct types in private
   attributes.
6. Tests: domain pure + use case with fake doubles + adapter integration (real docx).
7. Imports satisfy the hexagonal invariant: domain imports nothing from
   infrastructure; no local imports; no wildcard imports.
8. One class per file in the domain (exception: `domain/exceptions/` files may
   contain multiple exception classes).
9. Legacy system still functions end-to-end (coexistence verified).
