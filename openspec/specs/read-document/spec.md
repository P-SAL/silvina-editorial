# Read Document Specification

## Purpose

Read raw paragraph text from a `.docx` file through a port/adapter pair
(`DocumentTextPort` / python-docx adapter), exposed via `ReadDocumentUseCase`.
First infrastructure-backed capability in the hexagonal migration; faithfully
replicates legacy `WordReader.read_word_document()` behavior without
constructing `DocumentContentDTO` (that assembly belongs to a later slice).

## Requirements

### Requirement: DocumentTextPort Contract

`DocumentTextPort` MUST be an abstract base class at
`src/domain/document/document_text_port.py` (entity-folder placement, not a
top-level `domain/ports/` folder) exposing exactly one abstract method:
`read_paragraphs(path: str) -> list[str]`. The port MUST NOT import
`python-docx` or any other infrastructure library.

#### Scenario: Port defines a single abstract method

- GIVEN `DocumentTextPort`
- WHEN its interface is inspected
- THEN it declares exactly one abstract method, `read_paragraphs(path: str) -> list[str]`

#### Scenario: Port has zero infrastructure imports

- GIVEN the file defining `DocumentTextPort`
- WHEN its import statements are inspected
- THEN none import `docx` or anything from `src/infrastructure/`

### Requirement: Adapter Reads Non-Empty Stripped Paragraphs

The python-docx adapter implementing `DocumentTextPort` MUST open the file at
`path`, iterate its paragraphs, strip each paragraph's text, and return only
the non-empty stripped strings, in document order — byte-for-byte parity with
legacy `WordReader.read_word_document()`.

#### Scenario: Whitespace-only paragraphs are excluded

- GIVEN a `.docx` file containing paragraphs `["Title", "   ", "Body text"]`
- WHEN `read_paragraphs(path)` is called
- THEN the result is `["Title", "Body text"]`

#### Scenario: Leading and trailing whitespace is stripped

- GIVEN a `.docx` file containing a paragraph `"  Indented text  "`
- WHEN `read_paragraphs(path)` is called
- THEN the corresponding entry in the result is `"Indented text"`

#### Scenario: Paragraph order is preserved

- GIVEN a `.docx` file with paragraphs in a specific order
- WHEN `read_paragraphs(path)` is called
- THEN the returned list preserves that same order

#### Scenario: Document with no non-empty paragraphs returns an empty list

- GIVEN a `.docx` file whose only paragraphs are empty or whitespace-only
- WHEN `read_paragraphs(path)` is called
- THEN the result is `[]`

### Requirement: Adapter Raises Typed Exceptions at the I/O Boundary

The adapter MUST raise `DocumentNotFound` (from
`src.domain.exceptions.document_errors`) when the file at `path` does not
exist, and `DocumentUnreadable` when `python-docx` fails to open or parse an
existing file. Neither failure MUST propagate as a bare built-in exception
(e.g. `FileNotFoundError`, `PackageNotFoundError`) to the caller.

#### Scenario: Missing file raises DocumentNotFound

- GIVEN a `path` that does not point to an existing file
- WHEN `read_paragraphs(path)` is called
- THEN `DocumentNotFound` is raised

#### Scenario: Corrupt or unparseable file raises DocumentUnreadable

- GIVEN a `path` pointing to an existing file that `python-docx` cannot parse
  (e.g. not a valid `.docx` package)
- WHEN `read_paragraphs(path)` is called
- THEN `DocumentUnreadable` is raised

#### Scenario: Valid file raises neither exception

- GIVEN a `path` pointing to a well-formed `.docx` file
- WHEN `read_paragraphs(path)` is called
- THEN no exception is raised and a `list[str]` is returned

### Requirement: DocumentTextPort Is Consumed Directly by the Orchestrator

> **Superseded (2026-07-04, `refactor_analyze_document_wiring`)**: `ReadDocumentUseCase`
> and `ReadDocumentUseCaseWiring` were eliminated as redundant pass-through layers.
> `AnalyzeDocumentUseCase` now depends on `DocumentTextPort` directly and calls
> `read_paragraphs(path=document_path)` from its `execute()` method — see
> `openspec/specs/analyze-document/spec.md`, "Requirement: AnalyzeDocumentUseCase
> Orchestrator". `AnalyzeDocumentUseCaseWiring._get_document_text_port()` constructs
> the adapter directly (no intermediate sub-wiring).

`AnalyzeDocumentUseCase` MUST depend only on `DocumentTextPort` for reading paragraphs,
never constructing `DocumentContentDTO` from this step or adding business logic beyond
calling `read_paragraphs(path)` and using its result unchanged.

#### Scenario: Orchestrator uses the port's result unchanged

- GIVEN a `DocumentTextPort` test double returning `["A", "B"]` for a given path
- WHEN `AnalyzeDocumentUseCase.execute(path)` reads paragraphs via that port
- THEN the paragraphs used downstream are exactly `["A", "B"]`

#### Scenario: Port exceptions propagate unmodified through the orchestrator

- GIVEN a `DocumentTextPort` test double that raises `DocumentNotFound` for a
  given path
- WHEN `AnalyzeDocumentUseCase.execute(path)` is called with that path
- THEN `DocumentNotFound` propagates out of `execute()` unmodified (via the
  orchestrator's `@generic_error_handler`)

### Requirement: Behavioral Parity with Legacy WordReader

For any sample `.docx` document readable by both implementations, the
`DocumentTextPort` adapter's `read_paragraphs(path)` MUST return a list equal to
legacy `WordReader.read_word_document(path)`'s return value.

#### Scenario: Parity smoke test passes against a real sample document

- GIVEN a real sample `.docx` file used elsewhere in the test suite
- WHEN both legacy `WordReader.read_word_document(path)` and
  `AnalyzeDocumentUseCaseWiring()._get_document_text_port().read_paragraphs(path=path)`
  are called with that file
- THEN the two returned lists are equal element-for-element

## Out of Scope

- `win32com`-based counting (`data_access/word_counter.py`) — unrelated, Slice 6.
- Constructing `DocumentContentDTO` — Slice 6 (`ExtractContentUseCase`).
- Wiring `ReadDocumentUseCase` into `main.py` or `gradio_app.py` — Slice 14.
- `WordReader.read_document_with_styles()` and `get_document_properties()` —
  no current callers, not ported.
- Modifying or deleting `data_access/word_reader.py` — legacy stays untouched.
