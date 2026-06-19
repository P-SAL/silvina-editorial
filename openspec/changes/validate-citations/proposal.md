# Proposal: validate-citations (Slice 4)

## Intent

`business_logic/citation_matcher.py` (`CitationMatcher`, 186 lines) is pure computation —
regex/string matching only, no I/O, no LLM calls — sitting between `QualityAnalyzer` and
`StructureValidator` in `main.py`'s pipeline. `main.py:241-244` calls
`match_citations_to_references(section_type)` and consumes its result (`matched_count`,
`total_citations`, `unmatched_count`) to compute and print a match rate. This is the only
caller in the codebase. The migration lifts the matching logic that main.py actually uses
into the domain layer as a stateless service with a use case and wiring, following the
Slice 2/3 pattern exactly, while explicitly excluding presentation logic and confirmed dead
code that would otherwise inflate the slice.

## Scope

### In Scope

- `src/domain/citation/citation_matcher.py` — stateless `CitationMatcher` domain service:
  `find_orphaned_citations()`, `find_orphaned_references()`, `match_citations_to_references()`,
  and the private `_normalize_author()` helper (author-normalization regex rules preserved
  verbatim — non-author skip patterns, year/et-al/initials stripping).
- `src/application/match_citations_use_case.py` — `MatchCitationsUseCase.execute(citations:
  list[CitationDTO], references: list[ReferenceDTO], section_type: SectionName) ->
  CitationAnalysisResultDTO`.
- `src/infrastructure/wirings/match_citations_use_case_wiring.py` — `MatchCitationsUseCaseWiring`
  with `_get_*` accessor pattern (Slice 2/3 instance-based pattern).
- Hardcoded-string fix: `match_citations_to_references(section_type: str = "Referencias")`
  takes `section_type: SectionName` typed against the existing `SectionName` enum
  (`src/domain/enums/section_name.py`, `REFERENCES = "Referencias"`) instead of a literal string.
- Domain tests under `src/domain/tests/citation/test_citation_matcher.py` covering author
  normalization edge cases, orphaned citations, orphaned references, and the combined result.

### Out of Scope

- **`generate_report()`** — mixes Spanish presentation text (`"INTEGRIDAD DE CITAS..."`,
  emoji, severity labels) into business logic. Deferred to a future formatter adapter, same
  rationale as Slice 3's `apa_validator.generate_report()` exclusion.
- **`extract_all_citations()`** — confirmed dead code. Verified zero call sites anywhere in
  the codebase, including internally within `citation_matcher.py` itself (no other method
  calls it). Excluded explicitly, not silently dropped.
- **`business_logic/article_analyzer.py` (`ArticleAnalyzer`) — entire module out of scope.**
  Verified zero import/instantiation sites anywhere in the codebase (not wired into `main.py`
  or `gradio_app.py`). It is the only caller of `generate_report()`. Confirmed dead code.
- Deleting `business_logic/citation_matcher.py` — coexistence maintained until the caller
  switchover slice.
- Wiring `MatchCitationsUseCase` into `main.py` — deferred to the caller-switchover slice.
- **Orphaned-reference severity-by-section rule** (`generate_report()` lines 145-153: WARNING
  if `section_type == "Referencias"`, else INFO) is left untouched inside the legacy
  `generate_report()`, which has zero live callers (see above). Not extracted into a domain
  policy now — no current consumer needs it, and extracting it would be a premature
  abstraction. If a future slice needs this rule, lift it into a small pure function under
  `src/domain/citation/` (e.g. `severity_for_orphaned_references(section: SectionName) ->
  SeverityLevel`), not into a DTO — DTOs in this project are plain dataclasses with no
  behavior.

## Capabilities

### New Capabilities

- `validate-citations`: citation-reference integrity matching as a domain service — stateless
  computation of orphaned citations, orphaned references, and aggregate match statistics,
  exposed via a use case returning `CitationAnalysisResultDTO`.

### Modified Capabilities

None.

## Approach

1. Create `src/domain/citation/citation_matcher.py` (citation/ folder already exists from
   Slice 3's `apa_validator.py`).
2. Reuse existing DTOs verbatim — no new DTOs needed: `CitationDTO`, `ReferenceDTO`, and
   `CitationAnalysisResultDTO` already map field-for-field onto `CitationMatcher`'s actual
   inputs/outputs as consumed by `main.py`.
3. Make the service stateless per call: `find_orphaned_citations()` and
   `find_orphaned_references()` take `citations`/`references` as method parameters instead of
   relying on constructor-built `self.citation_keys`/`self.reference_keys` state, matching the
   Slice 3 precedent of dropping `self.violations`-style state.
4. `match_citations_to_references()` keeps its current logic (build valid-citations count,
   orphaned citations, return populated `CitationAnalysisResultDTO`) but takes `section_type:
   SectionName` instead of `str`.
5. `MatchCitationsUseCase.execute()` is a thin pass-through to the domain service.
6. `MatchCitationsUseCaseWiring.get_match_citations_use_case()` assembles the use case via
   `_get_*` accessors (no ports/adapters — pure computation, same as Slices 2-3).
7. Tests as `unittest.TestCase` under `src/domain/tests/citation/`, covering each non-author
   skip pattern (institutional acronyms, arXiv, DOI, date ranges) and the orphan-detection
   logic.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/domain/citation/citation_matcher.py` | New | Stateless `CitationMatcher` service |
| `src/application/match_citations_use_case.py` | New | `MatchCitationsUseCase` |
| `src/infrastructure/wirings/match_citations_use_case_wiring.py` | New | `MatchCitationsUseCaseWiring` |
| `src/domain/tests/citation/test_citation_matcher.py` | New | Domain service unit tests |
| `business_logic/citation_matcher.py` | Unchanged | Legacy stays alive during coexistence |
| `business_logic/article_analyzer.py` | Unchanged | Confirmed dead code, excluded from migration |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Author-normalization regex (non-author skip patterns, et-al/initials stripping) is a business rule that must be preserved exactly | Med | Copy regex verbatim; one test per skip pattern |
| Statelessness refactor (dropping constructor-built lookup dicts) could change behavior if a future caller relied on incremental matching | Low | `main.py`'s only call site uses the full result in a single call; no incremental usage exists today |
| `SectionType` enum exists alongside `SectionName` — risk of using the wrong enum | Low | `SectionName.REFERENCES = "Referencias"` is the literal match; explicitly use `SectionName`, not `SectionType` |

## Rollback Plan

All new files are additive. Legacy `business_logic/citation_matcher.py` is untouched. To roll
back: delete the 4 new source/test files. No existing behavior changes. `main.py` continues
importing from `business_logic/`. No migration state to undo.

## Dependencies

- Slice 2 (`validate-structure`) and Slice 3 (`validate-apa`) archived — establish the
  enum -> DTO -> domain service -> use case -> wiring pattern and the `src/domain/citation/`
  folder.
- `CitationDTO`, `ReferenceDTO`, `CitationAnalysisResultDTO` — already exist, reused as-is.
- `SectionName` enum — already exists, reused for the hardcoded-string fix.

## Success Criteria

- [ ] `CitationMatcher` domain service has no constructor-built state; all methods take their
      inputs as parameters
- [ ] `find_orphaned_citations()` and `find_orphaned_references()` preserve the exact author-
      normalization regex rules from the legacy class
- [ ] `match_citations_to_references()` accepts `section_type: SectionName`, not a raw string
- [ ] `MatchCitationsUseCase.execute()` returns `CitationAnalysisResultDTO` matching the fields
      `main.py:241-244` currently consumes (`total_citations`, `matched_count`,
      `unmatched_count`)
- [ ] `generate_report()`, `extract_all_citations()`, and `ArticleAnalyzer` are absent from the
      domain layer, with their exclusion documented as dead/presentation code
- [ ] Legacy `business_logic/citation_matcher.py` is unmodified; `main.py` still imports from
      `business_logic/`

## Open Questions

None — exploration's "dead integration point" concern is resolved: `main.py:241-244` actively
consumes `match_citations_to_references()`'s return value (match rate calculation), so the use
case must return a real, populated `CitationAnalysisResultDTO`, not a discarded one.
