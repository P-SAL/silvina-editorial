# Tasks: validate-citations (Slice 4)

Ordered, actionable checklist. Follows the Slice 2/3 phased structure: SCAFFOLD (none needed —
`src/domain/citation/` and `src/domain/tests/citation/` already exist from Slice 3) then RED/GREEN
TDD cycles per component, then full-suite regression.

Test runner: `python -m pytest src/` — baseline 209 tests passing.

## Phase 0 — SCAFFOLD

- [ ] **T-01** Confirm `src/domain/citation/__init__.py` and `src/domain/tests/citation/__init__.py`
  already exist (Slice 3) — no new folders required. No code change, verification only.
  - Satisfies: Module Layout (design.md)
  - Parallel: yes (can run alongside T-02 prep reading)

## Phase 1 — `CitationMatcher` domain service (RED -> GREEN per group)

Test file: `src/domain/tests/citation/test_citation_matcher.py`
Source file: `src/domain/citation/citation_matcher.py`

- [ ] **T-02** RED: write tests for `_normalize_author` non-author sentinel cases — one test per
  `_NON_AUTHOR_PATTERNS` entry: institutional acronym (`"UNESCO 2020"`), `arXiv:` identifier,
  `doi:` identifier, `repositorio` prefix, `"no hay"` prefix, multi-year date range
  (`"Datos 2018-2022"`). Each asserts `_normalize_author(...) == "__non_author__"`.
  - Satisfies: Requirement "Author Normalization Preserved Verbatim" (spec.md, scenarios:
    institutional acronym, arXiv, DOI, multi-year date range)
  - Sequential: must precede T-03 (no implementation exists yet)

- [ ] **T-03** GREEN: implement `CitationMatcher` class skeleton with `_NON_AUTHOR_PATTERNS` class
  constant and `_normalize_author` method (non-author branch only) to pass T-02. Copy regex list
  verbatim from design.md / legacy `business_logic/citation_matcher.py` lines 61-97.
  - Satisfies: Requirement "Author Normalization Preserved Verbatim"
  - Sequential: depends on T-02

- [ ] **T-04** RED: write tests for `_normalize_author` real-author normalization edge cases —
  et-al stripping (`"Wei, J. et al. (2022)"` -> `"wei"`), Spanish `" y "` co-author tail
  stripping, single-letter initials stripping, punctuation stripping, year-parenthetical
  stripping (including the `"de"`-date long form), case-folding to lowercase first word.
  - Satisfies: Requirement "Author Normalization Preserved Verbatim" (scenario: et-al and
    initials stripped)
  - Parallel: can be written alongside T-02 (independent test cases in the same file); must land
    before T-05

- [ ] **T-05** GREEN: complete `_normalize_author` body (year-parenthetical strip, et-al strip,
  `" y "` strip, initials strip, punctuation strip, lowercase first word) to pass T-04.
  - Satisfies: Requirement "Author Normalization Preserved Verbatim"
  - Sequential: depends on T-04, builds on T-03

- [ ] **T-06** RED: write tests for `_citable()` helper — given a mixed list of footnote
  citations, authorless citations, and valid authored non-footnote citations, assert `_citable()`
  returns only the valid authored non-footnote ones.
  - Satisfies: Requirement "Stateless Citation Matching Service" (supports orphan-detection and
    match-statistics requirements that depend on this filter)
  - Sequential: must precede T-07

- [ ] **T-07** GREEN: implement `_citable()` private helper (`citation_type != CitationType.FOOTNOTE
  and citation.author` filter) to pass T-06.
  - Satisfies: same as T-06
  - Sequential: depends on T-06

- [ ] **T-08** RED: write tests for `_build_citation_keys` and `_build_reference_keys` — assert
  each returns a `dict[str, DTO]` keyed by normalized author, built only from `_citable()`-filtered
  citations (for `_build_citation_keys`) or all references (for `_build_reference_keys`).
  - Satisfies: Requirement "Stateless Citation Matching Service" (no constructor state — keys
    built per call)
  - Sequential: depends on T-07 (needs `_citable()`)

- [ ] **T-09** GREEN: implement `_build_citation_keys` and `_build_reference_keys` to pass T-08.
  - Sequential: depends on T-08

- [ ] **T-10** RED: write tests for `_is_orphaned_citation` — citation whose normalized author is
  `"__non_author__"` is never orphaned; citation whose normalized author is absent from
  `reference_keys` is orphaned; citation whose normalized author is present is not orphaned.
  - Satisfies: Requirement "Orphaned Citation Detection" (non-author exclusion clause)
  - Sequential: depends on T-09 (needs `reference_keys` shape)

- [ ] **T-11** GREEN: implement `_is_orphaned_citation` to pass T-10.
  - Sequential: depends on T-10

- [ ] **T-12** RED: write tests for `find_orphaned_citations` — empty citations/references yields
  empty list; all citations matched yields empty list; one matched + one unmatched yields only
  the unmatched one; footnote citations excluded even when unmatched.
  - Satisfies: Requirement "Orphaned Citation Detection" (all four scenarios)
  - Sequential: depends on T-11

- [ ] **T-13** GREEN: implement `find_orphaned_citations` (delegates to `_citable`,
  `_build_reference_keys`, `_is_orphaned_citation`) to pass T-12.
  - Sequential: depends on T-12

- [ ] **T-14** RED: write tests for `find_orphaned_references` — all references cited yields empty
  list; one cited + one never-cited reference yields only the uncited one.
  - Satisfies: Requirement "Orphaned Reference Detection" (both scenarios)
  - Parallel: can be written alongside T-12 (independent method); must land before T-15

- [ ] **T-15** GREEN: implement `find_orphaned_references` (delegates to `_build_citation_keys`)
  to pass T-14.
  - Sequential: depends on T-14, T-09

- [ ] **T-16** RED: write tests for `match_citations_to_references` — empty inputs yield
  `total_citations=0, matched_count=0, unmatched_count=0, unmatched_citations=[]`; three matched
  authored non-footnote citations yield `matched_count=3, unmatched_count=0`; two citations with
  one orphaned yield `matched_count=1, unmatched_count=1` and `unmatched_citations` contains the
  orphaned citation's text; one footnote + one authored citation yields `total_citations=1`
  (footnote excluded); verify `section_type` accepts `SectionName.REFERENCES` enum member (no
  raw-string comparison).
  - Satisfies: Requirement "Aggregate Citation-Reference Match Statistics" (all five scenarios)
    and Requirement "Typed Section Parameter"
  - Sequential: depends on T-13 (uses `find_orphaned_citations`)

- [ ] **T-17** GREEN: implement `match_citations_to_references` (delegates to
  `find_orphaned_citations`, `_citable`; builds `CitationAnalysisResultDTO`;
  `section_type: SectionName = SectionName.REFERENCES`) to pass T-16.
  - Satisfies: same as T-16
  - Sequential: depends on T-16

- [ ] **T-18** RED: write a statelessness regression test — instantiate one `CitationMatcher`,
  call `find_orphaned_citations` twice with two different, unrelated citation/reference lists, and
  assert the second call's result is unaffected by the first call's inputs (e.g., no leaked
  `__non_author__` keys or counts between calls).
  - Satisfies: Requirement "Stateless Citation Matching Service" (scenario: repeated calls produce
    consistent results without shared state)
  - Sequential: depends on T-17 (exercises the fully implemented class)

- [ ] **T-19** GREEN: confirm T-18 passes with no code change (statelessness is structural —
  method parameters only, no `__init__`). If it fails, fix any accidental instance-state leak
  before proceeding.
  - Sequential: depends on T-18

## Phase 2 — `MatchCitationsUseCase` (RED -> GREEN)

Test file: `src/application/tests/test_match_citations_use_case.py`
Source file: `src/application/match_citations_use_case.py`

- [ ] **T-20** RED: write test asserting `MatchCitationsUseCase.execute(citations, references,
  section_type)` returns the exact same `CitationAnalysisResultDTO` values that calling
  `CitationMatcher().match_citations_to_references(...)` directly would produce, for a
  representative non-trivial citations/references pair (mirrors
  `test_validate_apa_use_case.py`'s `test_s12`/`test_s13` style — `test_s2x` numbering).
  - Satisfies: Requirement "MatchCitationsUseCase Pass-Through"
  - Sequential: depends on T-19 (needs a working `CitationMatcher`)

- [ ] **T-21** GREEN: implement `MatchCitationsUseCase.__init__(self, matcher: CitationMatcher)`
  and `execute(...)` as a thin delegate to `matcher.match_citations_to_references(...)` — no
  business logic, no empty-list guard (unlike `ValidateApaUseCase`, not needed here per design.md).
  - Satisfies: Requirement "MatchCitationsUseCase Pass-Through"
  - Sequential: depends on T-20

## Phase 3 — `MatchCitationsUseCaseWiring` (RED -> GREEN)

Test file: `src/infrastructure/tests/test_match_citations_use_case_wiring.py`
Source file: `src/infrastructure/wirings/match_citations_use_case_wiring.py`

- [ ] **T-22** RED: write tests mirroring `test_validate_apa_wiring.py` — `create_use_case()`
  returns a `MatchCitationsUseCase` instance; calling `execute([], [], SectionName.REFERENCES)` on
  the wired use case returns a `CitationAnalysisResultDTO`.
  - Satisfies: Requirement "MatchCitationsUseCaseWiring Assembly" (scenario: wiring produces a
    usable use case instance)
  - Sequential: depends on T-21

- [ ] **T-23** GREEN: implement `MatchCitationsUseCaseWiring` with `create_use_case()` public
  method and `_get_citation_matcher()` private accessor, replicating `ValidateApaWiring` literally
  (no ports/adapters needed — pure computation).
  - Satisfies: Requirement "MatchCitationsUseCaseWiring Assembly"
  - Sequential: depends on T-22

## Phase 4 — Full-suite regression

- [ ] **T-24** Run `python -m pytest src/` and confirm 0 regressions: baseline 209 tests + all new
  tests from T-02 through T-23 pass (estimated +25 to +35 new test cases across the three test
  files). No skipped or xfailed tests introduced.
  - Satisfies: all requirements collectively (final acceptance gate)
  - Sequential: depends on all prior tasks

- [ ] **T-25** Quick manual review pass against `clean-architecture/SKILL.md`: confirm no
  abbreviations, no inline comments in production code, docstrings present on public
  classes/methods, import order (domain has zero application/infrastructure imports), file naming
  matches class naming (`citation_matcher.py`, `match_citations_use_case.py`,
  `match_citations_use_case_wiring.py`).
  - Sequential: depends on T-24 (or can run in parallel with T-24 as a static check)

## Explicitly Out of Scope (carried from spec.md / design.md — no task needed)

- `generate_report()`, `extract_all_citations()`, `ArticleAnalyzer` — confirmed dead code /
  presentation logic, not migrated.
- Orphaned-reference severity-by-section rule — left in legacy `generate_report()`, no task.
- Deleting `business_logic/citation_matcher.py` or switching `main.py` callers — deferred to a
  future caller-switchover slice, no task here.

## Review Workload Forecast

- **Estimated changed/added lines**: ~340-380 total
  - `src/domain/citation/citation_matcher.py`: ~95 lines (new)
  - `src/application/match_citations_use_case.py`: ~20 lines (new)
  - `src/infrastructure/wirings/match_citations_use_case_wiring.py`: ~13 lines (new)
  - `src/domain/tests/citation/test_citation_matcher.py`: ~140-160 lines (new, ~20-25 test
    methods)
  - `src/application/tests/test_match_citations_use_case.py`: ~25 lines (new, 1-2 test methods)
  - `src/infrastructure/tests/test_match_citations_use_case_wiring.py`: ~20 lines (new, 2 test
    methods)
- **400-line budget risk**: Low. Estimated total sits comfortably under the 400-line PR budget,
  with roughly 20-60 lines of headroom even if test verbosity runs slightly over estimate. No
  production code modifications to existing files (purely additive — legacy
  `business_logic/citation_matcher.py` is untouched), which keeps the diff size predictable
  compared to Slice 2/3's DTO-touching changes.
- **Chained PRs recommended**: No. This slice is small enough and self-contained enough (3 new
  source files, 3 new test files, zero modifications to existing files) to ship as a single PR.
- **Decision needed before apply**: No. No ambiguity requiring a stop-and-ask — the design is
  fully specified with exact code in design.md, and the estimated size is well within budget.
