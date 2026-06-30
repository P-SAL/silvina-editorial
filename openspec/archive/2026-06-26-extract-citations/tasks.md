# Tasks: Extract Citations (Slice 7 — Hexagonal Migration)

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~500–600 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR-1 (domain) → PR-2 (infra adapters) → PR-3 (use case + wiring) |
| Delivery strategy | auto-chain |
| Chain strategy | stacked-to-main |

Decision needed before apply: No
Chained PRs recommended: Yes
Chain strategy: stacked-to-main
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Domain ports, exceptions, DTO, fakes | PR-1 | Base: refactor/hexagonal-migration |
| 2 | DocxCitationAdapter + DocxReferenceAdapter | PR-2 | Base: PR-1 branch |
| 3 | ExtractCitationsUseCase + wiring | PR-3 | Base: PR-2 branch |

---

## Phase 1: PR-1 — Domain Layer

- [ ] 1.1 RED: Create `src/domain/tests/document/test_citation_extraction_port.py` — S1a (TypeError on direct instantiation), S1b (signature check), S5a (fake returns configured list), S5b (fake raises configured error)
- [ ] 1.2 GREEN: Create `src/domain/document/citation_extraction_port.py` — `CitationExtractionPort(ABC)` with `@abstractmethod extract_citations(self, docx_path: str) -> list[CitationDTO]`
- [ ] 1.3 GREEN: Create `src/domain/tests/document/fake_citation_extraction_port.py` — `FakeCitationExtractionPort(CitationExtractionPort)` with configurable `citations` and `error`
- [ ] 1.4 RED: Create `src/domain/tests/document/test_reference_extraction_port.py` — S2a (TypeError), S2b (signature), S6a (fake returns tuple), S6b (fake raises error)
- [ ] 1.5 GREEN: Create `src/domain/document/reference_extraction_port.py` — `ReferenceExtractionPort(ABC)` with `@abstractmethod extract_references(self, docx_path: str) -> tuple[list[ReferenceDTO], str]`
- [ ] 1.6 GREEN: Create `src/domain/tests/document/fake_reference_extraction_port.py` — `FakeReferenceExtractionPort(ReferenceExtractionPort)` with configurable `result` tuple and `error`
- [ ] 1.7 RED: Create `src/domain/tests/exceptions/test_reference_errors.py` — S3a (MRO: ReferenceParsingFailed → ReferenceError → BaseSrcError → Exception), S3b (non-empty MESSAGE)
- [ ] 1.8 GREEN: Create `src/domain/exceptions/reference_errors.py` — `ReferenceError(BaseSrcError)` and `ReferenceParsingFailed(ReferenceError)` with non-empty `MESSAGE`
- [ ] 1.9 RED: Create `src/domain/tests/dtos/test_citation_extraction_result_dto.py` — S4a (FrozenInstanceError on any field assign), S4b (no field has a default value)
- [ ] 1.10 GREEN: Create `src/domain/dtos/citation_extraction_result_dto.py` — `@dataclass(frozen=True) CitationExtractionResultDTO(BaseDTO)` with required fields `citations: list[CitationDTO]`, `references: list[ReferenceDTO]`, `section_type: str`

---

## Phase 2: PR-2 — Infrastructure Adapters

- [ ] 2.1 RED: Create `src/infrastructure/tests/adapters/document/test_docx_citation_adapter.py` — S7a (non-empty `list[CitationDTO]` from `1. test_Científico.docx`), S7b (all items are CitationDTO), S7c (all `citation_type == CitationType.AUTHOR_YEAR`)
- [ ] 2.2 GREEN: Create `src/infrastructure/adapters/document/docx_citation_adapter.py` — `DocxCitationAdapter(CitationExtractionPort)`; zipfile + ElementTree parsing; three regex passes in order: parenthetical → multi-author narrative → single-author narrative; `@generic_error_handler`; raises `CitationParsingFailed` on failure; no `data_access` imports
- [ ] 2.3 RED: Create `src/infrastructure/tests/adapters/document/test_docx_reference_adapter.py` — S8a (non-empty `(list[ReferenceDTO], str)` from fixture), S8b (all items are ReferenceDTO), S8c (`section_type` in `{"Bibliografía", "Referencias", "Fuentes bibliográficas"}`)
- [ ] 2.4 GREEN: Create `src/infrastructure/adapters/document/docx_reference_adapter.py` — `DocxReferenceAdapter(ReferenceExtractionPort)`; zipfile + raw regex on XML string; year-end split for `_parse_references`; `@generic_error_handler`; raises `ReferenceParsingFailed` on failure; no `data_access` imports

---

## Phase 3: PR-3 — Use Case + Wiring

- [ ] 3.1 RED: Create `src/application/tests/test_extract_citations_use_case.py` — S9a (returns `CitationExtractionResultDTO` with fake ports), S9b (`.citations`, `.references`, `.section_type` match fake-injected values)
- [ ] 3.2 GREEN: Create `src/application/extract_citations_use_case.py` — `ExtractCitationsUseCase(citation_port, reference_port)` storing as `_citation_port` and `_reference_port`; `@generic_error_handler` on `execute(self, docx_path: str) -> CitationExtractionResultDTO`
- [ ] 3.3 RED: Create `src/infrastructure/tests/test_extract_citations_use_case_wiring.py` — S10a (`isinstance(uc._citation_port, DocxCitationAdapter)` and `isinstance(uc._reference_port, DocxReferenceAdapter)` both True)
- [ ] 3.4 GREEN: Create `src/infrastructure/wirings/extract_citations_use_case_wiring.py` — `ExtractCitationsUseCaseWiring.create_use_case()` wires `DocxCitationAdapter()` → `_citation_port` and `DocxReferenceAdapter()` → `_reference_port`
