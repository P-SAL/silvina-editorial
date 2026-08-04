# Citation Integrity Matching Specification

## Purpose

Stateless domain computation that matches in-text citations against bibliographic
references, surfacing orphaned citations (cited but not referenced) and orphaned
references (listed but never cited), plus aggregate match statistics consumed by
the document analysis pipeline.

## Requirements

### Requirement: Stateless Citation Matching Service

`CitationMatcher` MUST be a stateless domain service with no constructor-built
state. All matching methods MUST accept `citations: list[CitationDTO]` and
`references: list[ReferenceDTO]` as method parameters, not as instance attributes.

#### Scenario: Repeated calls produce consistent results without shared state

- GIVEN two separate calls to `find_orphaned_citations()` with different citation lists
- WHEN both calls execute on the same `CitationMatcher` instance
- THEN each call's result depends only on its own parameters, not on prior calls

### Requirement: Author Normalization Preserved Verbatim

`_normalize_author(text: str) -> str` MUST preserve the legacy regex rules exactly:
skip non-author patterns (institutional acronyms, `arXiv:`, `doi:`, `repositorio`,
`no hay`, multi-year date ranges) returning `"__non_author__"`; otherwise strip the
year parenthetical, `et al.` tail, Spanish `" y "` co-author tail, single-letter
initials, and punctuation, then return the lowercased first remaining word.

#### Scenario: Institutional acronym is treated as non-author

- GIVEN text `"UNESCO 2020"`
- WHEN `_normalize_author` is called
- THEN it returns `"__non_author__"`

#### Scenario: arXiv identifier is treated as non-author

- GIVEN text `"arXiv:2404.19573"`
- WHEN `_normalize_author` is called
- THEN it returns `"__non_author__"`

#### Scenario: DOI identifier is treated as non-author

- GIVEN text `"doi:10.1234/example"`
- WHEN `_normalize_author` is called
- THEN it returns `"__non_author__"`

#### Scenario: Multi-year date range is treated as non-author

- GIVEN text containing two distinct four-digit years (e.g. `"Datos 2018-2022"`)
- WHEN `_normalize_author` is called
- THEN it returns `"__non_author__"`

#### Scenario: Et-al and initials are stripped from a real author

- GIVEN text `"Wei, J. et al. (2022)"`
- WHEN `_normalize_author` is called
- THEN it returns `"wei"`

### Requirement: Orphaned Citation Detection

`find_orphaned_citations(citations, references) -> list[CitationDTO]` MUST return
citations whose normalized author has no matching entry among normalized
references. Footnote citations (`CitationType.FOOTNOTE`) and citations without an
`author` MUST be excluded from consideration. Citations whose normalized author is
`"__non_author__"` MUST NOT be reported as orphaned.

#### Scenario: No citations and no references yields no orphans

- GIVEN an empty citations list and an empty references list
- WHEN `find_orphaned_citations` is called
- THEN it returns an empty list

#### Scenario: All citations have matching references

- GIVEN citations normalizing to authors that all appear among the references
- WHEN `find_orphaned_citations` is called
- THEN it returns an empty list

#### Scenario: Some citations lack a matching reference

- GIVEN one citation normalizing to an author present in references and one
  normalizing to an author absent from references
- WHEN `find_orphaned_citations` is called
- THEN it returns only the citation with no matching reference

#### Scenario: Footnote citations are excluded from orphan detection

- GIVEN a citation with `citation_type = CitationType.FOOTNOTE` and no matching reference
- WHEN `find_orphaned_citations` is called
- THEN that citation is not included in the result

### Requirement: Orphaned Reference Detection

`find_orphaned_references(citations, references) -> list[ReferenceDTO]` MUST return
references whose normalized text has no matching entry among normalized,
non-footnote, authored citations.

#### Scenario: All references are cited

- GIVEN references normalizing to authors that all appear among the citations
- WHEN `find_orphaned_references` is called
- THEN it returns an empty list

#### Scenario: Some references are never cited

- GIVEN one reference normalizing to an author cited in text and one reference
  normalizing to an author never cited
- WHEN `find_orphaned_references` is called
- THEN it returns only the uncited reference

### Requirement: Aggregate Citation-Reference Match Statistics

`match_citations_to_references(citations, references, section_type: SectionName) ->
CitationAnalysisResultDTO` MUST count only non-footnote, authored citations as
`total_citations`, MUST set `matched_count` to `total_citations` minus the orphaned
citation count (floored at zero), MUST set `unmatched_count` to the orphaned
citation count, and MUST populate `unmatched_citations` with the orphaned citations'
text. `section_type` MUST be typed as `SectionName`, not `str`.

#### Scenario: No citations and no references

- GIVEN an empty citations list and an empty references list
- WHEN `match_citations_to_references` is called with `section_type=SectionName.REFERENCES`
- THEN the result has `total_citations=0`, `matched_count=0`, `unmatched_count=0`,
  and an empty `unmatched_citations`

#### Scenario: All citations matched

- GIVEN three authored, non-footnote citations all matching a reference
- WHEN `match_citations_to_references` is called
- THEN `matched_count` equals 3 and `unmatched_count` equals 0

#### Scenario: Some citations orphaned

- GIVEN two authored, non-footnote citations where one has no matching reference
- WHEN `match_citations_to_references` is called
- THEN `matched_count` equals 1, `unmatched_count` equals 1, and
  `unmatched_citations` contains the orphaned citation's text

#### Scenario: Footnote citations excluded from total_citations

- GIVEN one authored non-footnote citation and one footnote citation
- WHEN `match_citations_to_references` is called
- THEN `total_citations` equals 1, excluding the footnote citation

### Requirement: Typed Section Parameter

`match_citations_to_references` MUST accept `section_type: SectionName` (the
existing enum at `src/domain/enums/section_name.py`) instead of a raw `str` with a
hardcoded `"Referencias"` default.

#### Scenario: Caller passes the References enum member

- GIVEN `section_type=SectionName.REFERENCES`
- WHEN `match_citations_to_references` is called
- THEN the call succeeds with no string-literal comparison involved

### Requirement: CitationMatcher Is Consumed Directly by the Orchestrator

> **Superseded (2026-07-04, `refactor_analyze_document_wiring`)**: `MatchCitationsUseCase`
> and `MatchCitationsUseCaseWiring` were eliminated as redundant pass-through layers.
> `AnalyzeDocumentUseCase` now depends on `CitationMatcher` directly and calls
> `.match_citations_to_references(citations=..., references=..., section_type=...)`
> from its `execute()` method — see `openspec/specs/analyze-document/spec.md`.
> `AnalyzeDocumentUseCaseWiring._get_citation_matcher()` constructs the domain service
> directly (no intermediate sub-wiring).

`AnalyzeDocumentUseCase` MUST call `self._citation_matcher.match_citations_to_references(
citations=citations, references=references, section_type=section_name)` without adding
business logic of its own.

#### Scenario: Orchestrator uses the domain service's result unchanged

- GIVEN a citations list, a references list, and a section type
- WHEN `AnalyzeDocumentUseCase.execute()` calls `self._citation_matcher.match_citations_to_references(...)`
- THEN the returned `CitationAnalysisResultDTO` matches what
  `CitationMatcher.match_citations_to_references` would produce for the same inputs

## Out of Scope

- `generate_report()` — mixes Spanish presentation text, emoji, and severity
  labels into business logic; deferred to a future formatter adapter.
- `extract_all_citations()` — confirmed dead code (zero call sites anywhere,
  including internally); excluded explicitly.
- `business_logic/article_analyzer.py` (`ArticleAnalyzer`) — confirmed dead code,
  zero import/instantiation sites; entire module excluded.
- Orphaned-reference severity-by-section rule (`generate_report()` lines 145-153:
  WARNING when `section_type == "Referencias"`, else INFO) — left untouched inside
  legacy `generate_report()`, which has no live callers. Not extracted into a
  domain policy in this slice; if a future slice needs it, it MUST become a pure
  function under `src/domain/citation/` (e.g. `severity_for_orphaned_references
  (section: SectionName) -> SeverityLevel`), never logic inside a DTO.
- Deleting `business_logic/citation_matcher.py` or wiring the new use case into
  `main.py` — both deferred to the caller-switchover slice.
