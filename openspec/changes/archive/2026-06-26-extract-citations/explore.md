# Exploration: ExtractCitations (Slice 7 — Hexagonal Migration)

## Current State

Two legacy classes handle citation/reference document I/O:

- `data_access/citation_parser.py` — `CitationParser.extract_from_docx()` reads DOCX XML via `zipfile` + `xml.etree.ElementTree`, runs multi-pass regex over joined paragraph text, returns legacy `Citation` objects. Separate methods: `extract_footnotes(doc)` (takes a python-docx `Document` object), `parse(text, paragraph_index)` (per-paragraph fallback).
- `data_access/reference_parser.py` — `ReferenceParser.parse_from_docx()` reads DOCX XML via `zipfile` + raw regex (NOT ElementTree), detects bibliography section header, splits by year pattern, returns `(list[Reference], str)` where the string is one of `"Bibliografía"`, `"Referencias"`, or `"Fuentes bibliográficas"`.

The legacy `business_logic/citation_matcher.py` has `extract_all_citations()` which calls `extract_footnotes()` + `parse()` on a python-docx Document. This legacy class is NOT part of the hexagonal path.

The hexagonal `src/domain/citation/citation_matcher.py` receives `citations: list[CitationDTO]` and `references: list[ReferenceDTO]` as parameters — it does NO I/O and never calls `extract_footnotes`.

## Affected Areas

**Legacy (read-only, no changes):**
- `data_access/citation_parser.py` — logic source for `DocxCitationAdapter`
- `data_access/reference_parser.py` — logic source for `DocxReferenceAdapter`
- `business_logic/citation_matcher.py` — legacy only; not in `src/`

**New files (19 total):**

Domain ports (`src/domain/document/`):
- `src/domain/document/citation_extraction_port.py` — `CitationExtractionPort(ABC)`
- `src/domain/document/reference_extraction_port.py` — `ReferenceExtractionPort(ABC)`

Domain exceptions:
- `src/domain/exceptions/reference_errors.py` — `ReferenceError(BaseSrcError)`, `ReferenceParsingFailed(ReferenceError)`

Domain DTOs:
- `src/domain/dtos/citation_extraction_result_dto.py` — `CitationExtractionResultDTO(BaseDTO)` frozen dataclass

Adapters (`src/infrastructure/adapters/document/`):
- `src/infrastructure/adapters/document/docx_citation_adapter.py`
- `src/infrastructure/adapters/document/docx_reference_adapter.py`

Application:
- `src/application/extract_citations_use_case.py` — `ExtractCitationsUseCase`

Wiring:
- `src/infrastructure/wirings/extract_citations_use_case_wiring.py`

Tests — domain (7 files):
- `src/domain/tests/document/test_citation_extraction_port.py`
- `src/domain/tests/document/test_reference_extraction_port.py`
- `src/domain/tests/document/fake_citation_extraction_port.py`
- `src/domain/tests/document/fake_reference_extraction_port.py`
- `src/domain/tests/exceptions/test_reference_error.py`
- `src/domain/tests/exceptions/test_reference_parsing_failed.py`
- `src/domain/tests/dtos/test_citation_extraction_result.py`

Tests — infrastructure (4 files):
- `src/infrastructure/tests/adapters/document/test_docx_citation_adapter.py`
- `src/infrastructure/tests/adapters/document/test_docx_reference_adapter.py`
- `src/application/tests/test_extract_citations_use_case.py`
- `src/infrastructure/tests/test_extract_citations_use_case_wiring.py`

## Open Questions — Resolved

**1. ¿`extract_footnotes` pertenece a `CitationExtractionPort`?**
NO. Cero matches en `src/`. La versión hexagonal de `CitationMatcher` recibe DTOs directamente. El port expone un solo método: `extract_citations(docx_path: str) -> list[CitationDTO]`.

**2. ¿`SectionName` enum o `str` para section type?**
El port retorna `str`. `SectionName` solo tiene `REFERENCES = "Referencias"` — "Bibliografía" y "Fuentes bibliográficas" están ausentes. El layer de extracción mantiene `str` en todo el camino.

## Approaches Compared

| Approach | Pros | Cons |
|---|---|---|
| **A: Ports en `src/domain/document/`** (recommended) | Consistente con Slices 5 y 6 | — |
| B: Ports en `src/domain/ports/` | Menos directorios | Agrupación semántica incorrecta |
| C: Port combinado | Menos ABCs | Viola SRP |

## Constraints / Gotchas

1. Adapters importan de `src.domain.dtos`, NO de `domain.models` (legacy).
2. `ReferenceDTO.text` es el único campo poblado — el parser legacy no descompone `authors`, `year`, etc. Los demás quedan `None`. Deuda técnica documentada.
3. `@generic_error_handler` solo en métodos de adapter y `use_case.execute()`, no en ABCs.
4. `ReferenceParser.parse_from_docx()` usa regex raw sobre el XML string (no ElementTree). `CitationParser.extract_from_docx()` usa ElementTree con namespace dict. Portar esta lógica inline en los adapters — NO importar de `data_access/`.
5. Tests de integración del adapter necesitan `docs/sample-documents/1. test_Científico.docx` — verificar que tiene citas en texto Y sección de referencias.
6. `section_type` es uno de tres strings crudos: no coercionar a enum dentro del adapter ni del use case.

## Risks

- Portar lógica regex compleja sin tests primero → mitigado por strict TDD
- El documento de muestra debe tener citas en texto + sección de referencias — verificar antes de tests de integración del adapter
- `ReferenceDTO` parcialmente poblado crea deuda para slices futuros que necesiten campos estructurados
