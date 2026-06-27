# Design: Extract Citations (Slice 7 — Hexagonal Migration)

## Technical Approach

Port the XML-parsing logic of `data_access/citation_parser.py` and `data_access/reference_parser.py`
into the hexagonal layer: two ABCs in `src/domain/document/`, two adapters in
`src/infrastructure/adapters/document/`, two fake ports in `src/domain/tests/document/`,
one use case in `src/application/`, one wiring in `src/infrastructure/wirings/`, one result DTO,
and one exception hierarchy extension. Zero imports from `data_access/`.

## Architecture Decisions

| Decision | Choice | Rejected | Rationale |
|---|---|---|---|
| XML strategy — citations | `zipfile` + `xml.etree.ElementTree` with namespace | Raw regex on XML string | Correct paragraph boundaries; mirrors legacy `CitationParser.extract_from_docx` exactly |
| XML strategy — references | `zipfile` + raw regex on XML string | ElementTree | Simpler for full-document flat-text join; mirrors legacy `ReferenceParser.parse_from_docx` exactly |
| `section_type` type | `str` | `SectionName` enum | Spec explicitly forbids enum coercion; avoids vocabulary coupling |
| `@generic_error_handler` scope | Adapter public methods + use case `execute()` | ABC methods | Consistent with `DocxTextAdapter`, `ParagraphContentAdapter` |
| `CitationExtractionResultDTO` fields | All required, frozen | Optional fields with defaults | Spec R4; frozen matches existing DTO pattern |
| Wiring internal attributes | `_citation_port`, `_reference_port` | Other names | Wiring test S10a inspects these exact names; consistent with `ExtractContentUseCase` pattern |

## Data Flow

```
docx_path (str)
    │
    ├──► DocxCitationAdapter.extract_citations()
    │        ├── zipfile → word/document.xml → ET parse → paragraphs → full_text
    │        └── _extract_citations(full_text) → list[CitationDTO]
    │
    ├──► DocxReferenceAdapter.extract_references()
    │        ├── zipfile → word/document.xml → raw regex → bib_text + section_type
    │        └── _parse_references(bib_text) → list[ReferenceDTO]
    │
    └──► ExtractCitationsUseCase.execute(docx_path)
             └── CitationExtractionResultDTO(citations, references, section_type)
```

## D1 — Class Signatures

```python
# src/domain/document/citation_extraction_port.py
class CitationExtractionPort(ABC):
    @abstractmethod
    def extract_citations(self, docx_path: str) -> list[CitationDTO]: ...

# src/domain/document/reference_extraction_port.py
class ReferenceExtractionPort(ABC):
    @abstractmethod
    def extract_references(self, docx_path: str) -> tuple[list[ReferenceDTO], str]: ...

# src/domain/exceptions/reference_errors.py
class ReferenceError(BaseSrcError): ...
class ReferenceParsingFailed(ReferenceError):
    MESSAGE = "The reference could not be parsed."

# src/domain/dtos/citation_extraction_result_dto.py
@dataclass(frozen=True)
class CitationExtractionResultDTO(BaseDTO):
    citations: list[CitationDTO]
    references: list[ReferenceDTO]
    section_type: str

# src/domain/tests/document/fake_citation_extraction_port.py
class FakeCitationExtractionPort(CitationExtractionPort):
    def __init__(
        self,
        citations: list[CitationDTO] | None = None,
        error: Exception | None = None,
    ) -> None: ...
    def extract_citations(self, docx_path: str) -> list[CitationDTO]: ...

# src/domain/tests/document/fake_reference_extraction_port.py
class FakeReferenceExtractionPort(ReferenceExtractionPort):
    def __init__(
        self,
        result: tuple[list[ReferenceDTO], str] | None = None,
        error: Exception | None = None,
    ) -> None: ...
    def extract_references(self, docx_path: str) -> tuple[list[ReferenceDTO], str]: ...

# src/infrastructure/adapters/document/docx_citation_adapter.py
class DocxCitationAdapter(CitationExtractionPort):
    @generic_error_handler
    def extract_citations(self, docx_path: str) -> list[CitationDTO]: ...
    def _extract_citations(self, full_text: str) -> list[CitationDTO]: ...

# src/infrastructure/adapters/document/docx_reference_adapter.py
class DocxReferenceAdapter(ReferenceExtractionPort):
    @generic_error_handler
    def extract_references(self, docx_path: str) -> tuple[list[ReferenceDTO], str]: ...
    def _parse_references(self, bib_text: str) -> list[ReferenceDTO]: ...

# src/application/extract_citations_use_case.py
class ExtractCitationsUseCase:
    def __init__(
        self,
        citation_port: CitationExtractionPort,
        reference_port: ReferenceExtractionPort,
    ) -> None: ...
    @generic_error_handler
    def execute(self, docx_path: str) -> CitationExtractionResultDTO: ...

# src/infrastructure/wirings/extract_citations_use_case_wiring.py
class ExtractCitationsUseCaseWiring:
    def create_use_case(self) -> ExtractCitationsUseCase: ...
    def _get_citation_port(self) -> CitationExtractionPort: ...
    def _get_reference_port(self) -> ReferenceExtractionPort: ...
```

## D2 — Adapter Internals

### DocxCitationAdapter.extract_citations

```
1. with zipfile.ZipFile(docx_path, 'r') as zip_ref:
2.     doc_xml = zip_ref.read('word/document.xml').decode('utf-8')
3. root = ET.fromstring(doc_xml)
4. ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
5. paragraphs = []
6. for para in root.findall('.//w:p', ns):
7.     texts = [html.unescape(t.text) for t in para.findall('.//w:t', ns) if t.text]
8.     para_text = ''.join(texts)
9.     if para_text.strip(): paragraphs.append(para_text)
10. full_text = ' '.join(paragraphs)
11. return self._extract_citations(full_text)
```

### DocxCitationAdapter._extract_citations

State: `citations: list[CitationDTO]`, `seen: set[str]`, `multi_author_names: dict[str, set[str]]`,
`first_authors_by_year: dict[str, set[str]]`.

**Pass 1 — Parenthetical** (`re.findall(r'\([^)]*(?:19|20)\d{2}[^)]*\)', full_text)`):
- Skip if matches `r'^\(\d+\s+de\s+'` (date-like).
- Extract year via `r'(\d{4}[a-z]?)'`. Extract author as text before year, strip trailing comma.
- Skip if `len(author) < 2`. Deduplicate via `f"{author}|{year}"` key in `seen`.
- Append `CitationDTO(text=match, citation_type=CitationType.AUTHOR_YEAR, location=-1, author=author, year=year)`.

**Pre-pass**: for every multi-author parenthetical already collected (contains `'&'` or `','` in author),
extract first surname via `r'([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)'` and store in `first_authors_by_year[year]`.

**Pass 2 — Multi-author narrative**
(pattern: `r'(?<![a-záéíóúñ])\b([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ]+(?:\s+[ye&]\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\s]+?)+)\s+\((\d{4}[a-z]?)\)'`):
- Skip if `len(author) > 100` or starts with intro phrase (Como|Según|Si|No|En|El|La|Los|Las|Un|Una).
- Deduplicate. Collect each surname found in the author string into `multi_author_names[year]`.
- Append `CitationDTO(text=f"{author} ({year})", citation_type=CitationType.AUTHOR_YEAR, location=-1, author=author, year=year)`.

**Pass 3 — Single-author narrative**
(pattern: `r'(?<![(\[])\b([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)\s+\((\d{4}[a-z]?)\)'`):
- Skip if author in `first_authors_by_year.get(year, set())`.
- Skip if author in `multi_author_names.get(year, set())`.
- Skip if `f"{author}|{year}"` already in `seen`.
- Append `CitationDTO(text=f"{author} ({year})", citation_type=CitationType.AUTHOR_YEAR, location=-1, author=author, year=year)`.

Return `citations`.

### DocxReferenceAdapter.extract_references

```
1. with zipfile.ZipFile(docx_path, 'r') as zip_ref:
2.     doc_xml = zip_ref.read('word/document.xml').decode('utf-8')
3. text_pattern = r'<w:t[^>]*>([^<]+)</w:t>'
4. all_texts = [html.unescape(t) for t in re.findall(text_pattern, doc_xml)]
5. full_text_compact = ''.join(all_texts)
6. bib_match = re.search(
       r'(Bibliograf[íi]a|Referencias|Fuentes\s*bibliogr[áa]ficas(?:\s*consultadas)?)\s*(.{100,})',
       full_text_compact, re.IGNORECASE | re.DOTALL)
7. if not bib_match: return ([], "Referencias")
8. g1 = bib_match.group(1).lower()
9. if "ibliograf" in g1:   section_type = "Bibliografía"
   elif "fuentes" in g1:   section_type = "Fuentes bibliográficas"
   else:                   section_type = "Referencias"
10. bib_text = bib_match.group(2)
11. return (self._parse_references(bib_text), section_type)
```

### DocxReferenceAdapter._parse_references

```
year_end_pattern = r'\((?:\d{1,2}\s+de\s+\w+\s+de\s+)?\d{4}[a-z]?\)\.?'
parts = re.split(f'({year_end_pattern})', bib_text)
current_ref = ""
references = []

for part in parts:
    if re.match(year_end_pattern, part):
        current_ref += part
        current_ref = current_ref.strip()
        if len(current_ref) > 30:
            author_match = re.search(
                r'([A-ZÁ-ÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁ-ÚÑ]?[a-záéíóúñ]+)*,\s+[A-ZÁÉÍÓÚÑ]\..*)',
                current_ref)
            clean_ref = author_match.group(1).strip() if author_match else current_ref
            clean_ref = re.sub(r'^[-–—•]+\s*', '', clean_ref)
            references.append(ReferenceDTO(text=clean_ref))
        current_ref = ""
    else:
        current_ref += part

# Trailing fragment (same cleanup as above)
if current_ref.strip() and len(current_ref.strip()) > 30:
    ...append ReferenceDTO(text=clean_ref)

return references
```

## D3 — Fake Ports

```python
class FakeCitationExtractionPort(CitationExtractionPort):
    def __init__(
        self,
        citations: list[CitationDTO] | None = None,
        error: Exception | None = None,
    ) -> None:
        self._citations = citations if citations is not None else []
        self._error = error

    def extract_citations(self, docx_path: str) -> list[CitationDTO]:
        if self._error is not None:
            raise self._error
        return self._citations
```

```python
class FakeReferenceExtractionPort(ReferenceExtractionPort):
    def __init__(
        self,
        result: tuple[list[ReferenceDTO], str] | None = None,
        error: Exception | None = None,
    ) -> None:
        self._result = result if result is not None else ([], "Referencias")
        self._error = error

    def extract_references(self, docx_path: str) -> tuple[list[ReferenceDTO], str]:
        if self._error is not None:
            raise self._error
        return self._result
```

## D4 — Use Case Orchestration

Exact order of port calls in `execute(self, docx_path: str)`:
1. `citations = self._citation_port.extract_citations(docx_path)`
2. `references, section_type = self._reference_port.extract_references(docx_path)`
3. Return `CitationExtractionResultDTO(citations=citations, references=references, section_type=section_type)`

Constructor stores ports verbatim — no adaptation:
```python
def __init__(self, citation_port: CitationExtractionPort, reference_port: ReferenceExtractionPort) -> None:
    self._citation_port = citation_port
    self._reference_port = reference_port
```

## D5 — Wiring

```python
class ExtractCitationsUseCaseWiring:
    def create_use_case(self) -> ExtractCitationsUseCase:
        return ExtractCitationsUseCase(
            citation_port=self._get_citation_port(),
            reference_port=self._get_reference_port(),
        )

    def _get_citation_port(self) -> CitationExtractionPort:
        return DocxCitationAdapter()

    def _get_reference_port(self) -> ReferenceExtractionPort:
        return DocxReferenceAdapter()
```

Wiring test S10a inspects: `uc._citation_port` (must be `DocxCitationAdapter`) and
`uc._reference_port` (must be `DocxReferenceAdapter`).

## D6 — Import Map

| File | Imports |
|---|---|
| `citation_extraction_port.py` | `abc.ABC`, `abc.abstractmethod`, `src.domain.dtos.citation_dto.CitationDTO` |
| `reference_extraction_port.py` | `abc.ABC`, `abc.abstractmethod`, `src.domain.dtos.reference_dto.ReferenceDTO` |
| `reference_errors.py` | `src.domain.exceptions.base_src_error.BaseSrcError` |
| `citation_extraction_result_dto.py` | `dataclasses.dataclass`, `src.domain.dtos.base_dto.BaseDTO`, `src.domain.dtos.citation_dto.CitationDTO`, `src.domain.dtos.reference_dto.ReferenceDTO` |
| `fake_citation_extraction_port.py` | `src.domain.document.citation_extraction_port.CitationExtractionPort`, `src.domain.dtos.citation_dto.CitationDTO` |
| `fake_reference_extraction_port.py` | `src.domain.document.reference_extraction_port.ReferenceExtractionPort`, `src.domain.dtos.reference_dto.ReferenceDTO` |
| `docx_citation_adapter.py` | `html`, `re`, `zipfile`, `xml.etree.ElementTree as ET`, `src.domain.document.citation_extraction_port.CitationExtractionPort`, `src.domain.dtos.citation_dto.CitationDTO`, `src.domain.enums.citation_type.CitationType`, `src.domain.exceptions.citation_errors.CitationParsingFailed`, `src.domain.exceptions.decorators.generic_error_handler.generic_error_handler` |
| `docx_reference_adapter.py` | `html`, `re`, `zipfile`, `src.domain.document.reference_extraction_port.ReferenceExtractionPort`, `src.domain.dtos.reference_dto.ReferenceDTO`, `src.domain.exceptions.reference_errors.ReferenceParsingFailed`, `src.domain.exceptions.decorators.generic_error_handler.generic_error_handler` |
| `extract_citations_use_case.py` | `src.domain.document.citation_extraction_port.CitationExtractionPort`, `src.domain.document.reference_extraction_port.ReferenceExtractionPort`, `src.domain.dtos.citation_extraction_result_dto.CitationExtractionResultDTO`, `src.domain.exceptions.decorators.generic_error_handler.generic_error_handler` |
| `extract_citations_use_case_wiring.py` | `src.application.extract_citations_use_case.ExtractCitationsUseCase`, `src.domain.document.citation_extraction_port.CitationExtractionPort`, `src.domain.document.reference_extraction_port.ReferenceExtractionPort`, `src.infrastructure.adapters.document.docx_citation_adapter.DocxCitationAdapter`, `src.infrastructure.adapters.document.docx_reference_adapter.DocxReferenceAdapter` |

## File Changes

| File | Action | Description |
|---|---|---|
| `src/domain/document/citation_extraction_port.py` | Create | ABC port — R1 |
| `src/domain/document/reference_extraction_port.py` | Create | ABC port — R2 |
| `src/domain/exceptions/reference_errors.py` | Create | Exception hierarchy — R3 |
| `src/domain/dtos/citation_extraction_result_dto.py` | Create | Frozen result DTO — R4 |
| `src/domain/tests/document/fake_citation_extraction_port.py` | Create | Test double — R5 |
| `src/domain/tests/document/fake_reference_extraction_port.py` | Create | Test double — R6 |
| `src/infrastructure/adapters/document/docx_citation_adapter.py` | Create | XML adapter (ET) — R7 |
| `src/infrastructure/adapters/document/docx_reference_adapter.py` | Create | XML adapter (raw regex) — R8 |
| `src/application/extract_citations_use_case.py` | Create | Orchestrator use case — R9 |
| `src/infrastructure/wirings/extract_citations_use_case_wiring.py` | Create | DI factory — R10 |
| `src/domain/tests/document/test_citation_extraction_port.py` | Create | S1a, S1b, S5a, S5b |
| `src/domain/tests/document/test_reference_extraction_port.py` | Create | S2a, S2b, S6a, S6b |
| `src/domain/tests/exceptions/test_reference_errors.py` | Create | S3a, S3b |
| `src/domain/tests/dtos/test_citation_extraction_result_dto.py` | Create | S4a, S4b |
| `src/infrastructure/tests/adapters/document/test_docx_citation_adapter.py` | Create | S7a, S7b, S7c |
| `src/infrastructure/tests/adapters/document/test_docx_reference_adapter.py` | Create | S8a, S8b, S8c |
| `src/application/tests/test_extract_citations_use_case.py` | Create | S9a, S9b |
| `src/infrastructure/tests/test_extract_citations_use_case_wiring.py` | Create | S10a |

## Testing Strategy

| Layer | What | Approach |
|---|---|---|
| Domain/Port | ABC uninstantiable; signature matches spec | `pytest.raises(TypeError)`, `inspect.signature` |
| Domain/Exception | MRO chain; MESSAGE non-empty | `isinstance` chain, `assert ReferenceParsingFailed.MESSAGE` |
| Domain/DTO | Frozen; all fields required | `FrozenInstanceError`, `fields()` introspection |
| Domain/Fake | Configurable return + error raise | Unit; no I/O |
| Infra/Adapter | Real fixture `1. test_Científico.docx` | Integration — zipfile read; assert non-empty list + types |
| Application/UseCase | Fake ports injected | Unit — assert DTO fields delegate correctly |
| Infra/Wiring | `create_use_case()` returns correct types | `isinstance` on `_citation_port` and `_reference_port` |

## Migration / Rollout

No migration required. `data_access/` classes remain untouched. New hexagonal path operates in parallel until usage sites are migrated in a later slice.

## Open Questions

None — all design decisions are resolved by the spec and confirmed context.
