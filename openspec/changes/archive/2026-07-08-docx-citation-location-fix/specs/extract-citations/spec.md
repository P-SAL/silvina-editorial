# Delta for extract-citations

## ADDED Requirements

None

## MODIFIED Requirements

### Requirement: R7 — DocxCitationAdapter

`DocxCitationAdapter(CitationExtractionPort)` MUST exist at
`src/infrastructure/adapters/document/docx_citation_adapter.py`.
MUST accept an optional `max_author_name_length: int` in its constructor (defaulting to 100) and store it. During multi-author extraction, if the author name length exceeds `max_author_name_length`, it MUST NOT be parsed as an author.
The adapter MUST read the document paragraph by paragraph, tracking the 0-based paragraph index, and assign it to the `location` field of each extracted `CitationDTO`.
If citations are extracted via the legacy raw text fallback (`full_text` or `paragraphs` as `str`), it MUST assign `location=0` to the extracted `CitationDTO`s.
Constructs `CitationDTO` (not legacy `Citation`).
Raises `CitationParsingFailed` on failure. MUST NOT import from `data_access/`.
(Previously: All extracted citations were assigned location=-1.)

#### Scenario: S7a — Valid file

- GIVEN `1. test_Científico.docx`
- WHEN `extract_citations(docx_path)` is called
- THEN a non-empty `list[CitationDTO]` is returned without raising.

#### Scenario: S7b — Item type

- GIVEN result from S7a
- WHEN each item is inspected
- THEN all are `CitationDTO` instances.

#### Scenario: S7c — Citation type

- GIVEN result from S7a
- WHEN `citation_type` is accessed on any item
- THEN it is `CitationType.AUTHOR_YEAR`.

#### Scenario: S7d — Custom max_author_name_length

- GIVEN a `DocxCitationAdapter` initialized with `max_author_name_length` = 5
- WHEN a citation with a long author name is processed
- THEN it is rejected and not returned.

#### Scenario: S7e — Paragraph Index Assignment

- GIVEN a document with citations in multiple paragraphs
- WHEN `extract_citations(docx_path)` is called
- THEN each citation has its `location` set to its correct 0-based paragraph index.

#### Scenario: S7f — Raw Text Fallback

- GIVEN a raw text string or `full_text` fallback
- WHEN citations are extracted using `_extract_citations`
- THEN each citation has its `location` set to 0.

## REMOVED Requirements

None

## RENAMED Requirements

None
