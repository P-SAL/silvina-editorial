# Extract Citations — Specification

**Change**: extract-citations | **Slice**: 7 | **New files only**: 19

## Purpose

Close the hexagonal gap between citation-matching domain logic and the file system:
two ports, two adapters, one result DTO, one use case, one wiring — zero imports
from `data_access/`.

## Cross-Cutting Constraints

| Rule | Value |
|---|---|
| Import direction | Domain MUST NOT import `src.infrastructure.*` |
| Adapter imports | `src.domain.*` prefix (not bare `domain.*`) |
| Decorator import | `from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler` |
| One class per file | Domain layer (exception: `reference_errors.py` — two classes allowed) |
| Parameter name | `docx_path: str` everywhere (not `path`) |
| Sample fixture | `docs/sample-documents/1. test_Científico.docx` — verified present |
| Legacy coexistence | `data_access/` classes MUST remain untouched; system MUST work end-to-end |

## Requirements

### R1 — CitationExtractionPort

`CitationExtractionPort(ABC)` MUST exist at `src/domain/document/citation_extraction_port.py`.
Single method `extract_citations(docx_path: str) -> list[CitationDTO]` decorated with `@abstractmethod`.
No `@generic_error_handler` on the ABC method.

**S1a — Abstract**: GIVEN `CitationExtractionPort` | WHEN instantiated directly | THEN `TypeError` raised (cannot instantiate abstract class).
**S1b — Signature**: GIVEN port class | WHEN signature inspected | THEN param is `docx_path: str`, return annotation is `list[CitationDTO]`.

### R2 — ReferenceExtractionPort

`ReferenceExtractionPort(ABC)` MUST exist at `src/domain/document/reference_extraction_port.py`.
Single method `extract_references(docx_path: str) -> tuple[list[ReferenceDTO], str]` decorated with `@abstractmethod`.
No `@generic_error_handler` on the ABC method.

**S2a — Abstract**: GIVEN `ReferenceExtractionPort` | WHEN instantiated directly | THEN `TypeError` raised (cannot instantiate abstract class).
**S2b — Signature**: GIVEN port class | WHEN signature inspected | THEN param is `docx_path: str`, return annotation is `tuple[list[ReferenceDTO], str]`.

### R3 — ReferenceParsingFailed

`src/domain/exceptions/reference_errors.py` MUST define `ReferenceError(BaseSrcError)` and
`ReferenceParsingFailed(ReferenceError)`. `ReferenceParsingFailed.MESSAGE` MUST be a non-empty string.
Pattern follows `citation_errors.py` exactly.

**S3a — Inheritance**: GIVEN `ReferenceParsingFailed` | WHEN MRO checked | THEN chain is `ReferenceParsingFailed → ReferenceError → BaseSrcError → Exception`.
**S3b — MESSAGE**: GIVEN `ReferenceParsingFailed` | WHEN `.MESSAGE` accessed | THEN value is a non-empty string.

### R4 — CitationExtractionResultDTO

Frozen dataclass at `src/domain/dtos/citation_extraction_result_dto.py` inheriting `BaseDTO`.
Fields: `citations: list[CitationDTO]`, `references: list[ReferenceDTO]`, `section_type: str`.
No field has a default value.

**S4a — Frozen**: GIVEN an instance | WHEN any field is assigned | THEN `FrozenInstanceError` raised.
**S4b — All required**: GIVEN the dataclass | WHEN fields inspected | THEN none carry a default value.

### R5 — FakeCitationExtractionPort

`FakeCitationExtractionPort(CitationExtractionPort)` MUST exist at
`src/domain/tests/document/fake_citation_extraction_port.py`.
Configurable return value for `extract_citations` and optional error to raise.

**S5a — Return**: GIVEN fake configured with `[CitationDTO(...)]` | WHEN `extract_citations` called | THEN configured list returned.
**S5b — Error**: GIVEN fake configured with an exception | WHEN `extract_citations` called | THEN that exception raised.

### R6 — FakeReferenceExtractionPort

`FakeReferenceExtractionPort(ReferenceExtractionPort)` MUST exist at
`src/domain/tests/document/fake_reference_extraction_port.py`.
Configurable return tuple for `extract_references` and optional error.

**S6a — Return**: GIVEN fake configured with `([ReferenceDTO(...)], "Bibliografía")` | WHEN `extract_references` called | THEN configured tuple returned.
**S6b — Error**: GIVEN fake configured with an exception | WHEN `extract_references` called | THEN that exception raised.

### R7 — DocxCitationAdapter

`DocxCitationAdapter(CitationExtractionPort)` MUST exist at
`src/infrastructure/adapters/document/docx_citation_adapter.py`.
No `@generic_error_handler` on adapter methods; error wrapping is handled at the use case layer. Constructs `CitationDTO` (not legacy `Citation`).
Raises `CitationParsingFailed` on failure. MUST NOT import from `data_access/`.

**S7a — Valid file**: GIVEN `1. test_Científico.docx` | WHEN `extract_citations(docx_path)` called | THEN non-empty `list[CitationDTO]` returned without raising.
**S7b — Item type**: GIVEN result from S7a | WHEN each item inspected | THEN all are `CitationDTO` instances.
**S7c — Citation type**: GIVEN result from S7a | WHEN `citation_type` accessed on any item | THEN `CitationType.AUTHOR_YEAR`.

### R8 — DocxReferenceAdapter

`DocxReferenceAdapter(ReferenceExtractionPort)` MUST exist at
`src/infrastructure/adapters/document/docx_reference_adapter.py`.
No `@generic_error_handler` on adapter methods; error wrapping is handled at the use case layer.
`section_type` MUST be one of `{"Bibliografía", "Referencias", "Fuentes bibliográficas"}`.
Raises `ReferenceParsingFailed` on failure. MUST NOT import from `data_access/`.

**S8a — Valid file**: GIVEN `1. test_Científico.docx` | WHEN `extract_references(docx_path)` called | THEN `(non-empty list[ReferenceDTO], non-empty str)` returned without raising.
**S8b — Item type**: GIVEN returned list | WHEN items inspected | THEN all are `ReferenceDTO` instances.
**S8c — section_type**: GIVEN returned string | THEN value is one of the three allowed section names.

### R9 — ExtractCitationsUseCase

`ExtractCitationsUseCase` MUST exist at `src/application/extract_citations_use_case.py`.
Constructor: `(citation_port: CitationExtractionPort, reference_port: ReferenceExtractionPort)`.
`@generic_error_handler` on `execute(self, docx_path: str) -> CitationExtractionResultDTO`.

**S9a — Return type**: GIVEN fakes injected | WHEN `execute(docx_path)` called | THEN return is `CitationExtractionResultDTO`.
**S9b — Field delegation**: GIVEN fake citation port returns `[cit]` and fake reference port returns `([ref], "Bibliografía")` | WHEN `execute` called | THEN `.citations == [cit]`, `.references == [ref]`, `.section_type == "Bibliografía"`.

### R10 — ExtractCitationsUseCaseWiring

`ExtractCitationsUseCaseWiring` MUST exist at `src/infrastructure/wirings/extract_citations_use_case_wiring.py`.
`create_use_case() -> ExtractCitationsUseCase` MUST wire `DocxCitationAdapter` to `_citation_port`
and `DocxReferenceAdapter` to `_reference_port`.

**S10a — Correct types**: GIVEN `create_use_case()` called | WHEN private attributes inspected | THEN `isinstance(uc._citation_port, DocxCitationAdapter)` and `isinstance(uc._reference_port, DocxReferenceAdapter)` are both `True`.
