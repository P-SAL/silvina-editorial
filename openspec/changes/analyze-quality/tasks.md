# Tasks: analyze-quality (Slice 5) — PR-A Rework

> **Context**: PR-A is open as PR #13, NOT YET MERGED, with 273 passing tests against the
> original monolithic `QualityAnalyzer` (240 lines: sampling + prompt building + parsing +
> orchestration fused in one class). The user made 6 follow-up design decisions (ADR-6 through
> ADR-9 in `design.md`) before merging, splitting that monolith into focused collaborators. This
> tasks file replaces the original PR-A tasks file and governs the REWORK — it does not
> re-litigate the already-merged-in-spirit decisions (port shape, fallback constant, direct
> per-call assignment, full-call failure detection), only the new structural split.

Strangler/incremental approach: build the new collaborator classes first (still genuine TDD
RED/GREEN — these are new files even though their logic is ported, not invented), wire
`QualityAnalyzer` to delegate to them, delete the now-redundant monolithic code, redistribute the
9 existing test classes into 3 new test files, delete the old test file, then run a full
regression gate.

---

## Phase 0 — New domain primitives (parallel, independent files)

### T-01: RED — write `ReferenceLineMarker` enum test
- File: `src/domain/tests/enums/test_reference_line_marker.py`.
- `test_enum_has_exactly_four_members` — `len(ReferenceLineMarker) == 4`.
- `test_enum_members_have_expected_values` — `HTTP == "http"`, `DOI == "doi.org"`,
  `HTTPS == "https"`, `ISBN == "ISBN"`.
- **Satisfies**: Requirement "ReferenceLineMarker Enum Replaces the Reference-Line Tuple".
- **Parallel/Sequential**: Parallel with T-03, T-05. Sequential RED before T-02 GREEN.

### T-02: GREEN — implement `ReferenceLineMarker` enum
- File: `src/domain/enums/reference_line_marker.py` — exact code from design.md.
- **Satisfies**: same requirement.
- **Parallel/Sequential**: Sequential after T-01.

### T-03: RED — write `DimensionScoreDTO` test
- File: `src/domain/tests/dtos/test_dimension_score_dto.py`.
- `test_is_frozen_dataclass_extending_base_dto` — `issubclass(DimensionScoreDTO, BaseDTO)`,
  attempting attribute mutation raises `FrozenInstanceError` (mirrors existing DTO test pattern,
  e.g. `test_quality_result.py`).
- `test_holds_score_and_feedback_fields` — construct with `score=8.0, feedback="text"`, assert
  both fields readable.
- **Satisfies**: Requirement "DimensionScoreDTO and ParsedResponseDTO Extracted as Domain DTOs"
  (DTO half).
- **Parallel/Sequential**: Parallel with T-01, T-05. Sequential RED before T-04 GREEN.

### T-04: GREEN — implement `DimensionScoreDTO`
- File: `src/domain/dtos/dimension_score_dto.py` — exact code from design.md.
- **Satisfies**: same requirement.
- **Parallel/Sequential**: Sequential after T-03.

### T-05: RED — write `ParsedResponseDTO` test
- File: `src/domain/tests/dtos/test_parsed_response_dto.py`.
- `test_is_frozen_dataclass_extending_base_dto`.
- `test_default_scores_and_matched_dimensions_are_empty` — no-args construction yields `{}` and
  `frozenset()` (validates `field(default_factory=...)` wiring, not shared mutable defaults).
- `test_holds_scores_dict_and_matched_dimensions_frozenset` — construct with a populated dict
  keyed by `QualityDimension` and a non-empty `frozenset`.
- **Satisfies**: Requirement "DimensionScoreDTO and ParsedResponseDTO Extracted as Domain DTOs"
  (both scenarios — "extend BaseDTO" and "neither defined inside quality_analyzer.py", the latter
  verified later in T-19).
- **Parallel/Sequential**: Parallel with T-01, T-03. Depends on T-04 existing (imports
  `DimensionScoreDTO`) — sequence T-04 before T-05's GREEN, RED can be written referencing the
  not-yet-existing import per standard TDD.
- **Note**: requires `DimensionScoreDTO` to exist for GREEN (T-06) to pass; RED can be written in
  parallel but GREEN is sequential after T-04.

### T-06: GREEN — implement `ParsedResponseDTO`
- File: `src/domain/dtos/parsed_response_dto.py` — exact code from design.md.
- **Satisfies**: same requirement.
- **Parallel/Sequential**: Sequential after T-04 and T-05.

---

## Phase 1 — `QualityTextSampler` (new file, ported logic)

### T-07: RED — write `QualityTextSampler` tests (ported from old `test_quality_analyzer.py`)
- File: `src/domain/tests/quality/test_quality_text_sampler.py` — new file, calls
  `QualityTextSampler` directly, no fake `LlmGeneratorPort` needed.
- Port these existing TestCase classes/methods from the current
  `src/domain/tests/quality/test_quality_analyzer.py`, adapted to call `sampler.build_sample(...)`
  directly and assert on its return string instead of inspecting a captured prompt:
  - `TestTextSampling.test_short_document_uses_full_text_fallback_instead_of_sample`
  - `TestTextSampling.test_long_document_uses_strategic_sample_not_full_text`
  - `TestConclusionDetection.test_conclusion_paragraphs_exclude_reference_like_lines`
- Add 1 new test required by the updated spec (not present in the old suite, since the old
  sampler had no constructor parameters): `test_constructor_parameters_override_legacy_defaults`
  — construct `QualityTextSampler(min_sample_word_count=10, text_sample_character_limit=500)`,
  assert the fallback threshold and truncation length honor the constructor values, not
  `400`/`8000`.
- Add 1 new test for default-parameter equivalence:
  `test_defaults_match_legacy_hardcoded_constants` — construct with no arguments, assert behavior
  identical to the old `_MINIMUM_SAMPLE_WORD_COUNT=400`/`_TEXT_SAMPLE_CHARACTER_LIMIT=8000`.
- **Satisfies**: Requirement "QualityTextSampler Owns Text-Sampling Logic" (all 4 scenarios) and
  Requirement "Tunable Sampling Values Are Constructor Parameters, Not Domain Env Reads" (both
  scenarios, partially — zero-import-of-os/dotenv scenario verified in T-08/T-22).
- **Parallel/Sequential**: Sequential RED before T-08 GREEN. Can be written in parallel with
  Phase 0 and Phase 2's RED (T-09).

### T-08: GREEN — implement `QualityTextSampler`
- File: `src/domain/quality/quality_text_sampler.py` — exact code from design.md, ported verbatim
  from the current `_build_text_sample`/`_collect_conclusion_or_tail_paragraphs`/
  `_is_reference_like` in `quality_analyzer.py`, with:
  - Module constants (`_MINIMUM_SAMPLE_WORD_COUNT`, `_TEXT_SAMPLE_CHARACTER_LIMIT`) replaced by
    constructor parameters (`min_sample_word_count: int = 400`,
    `text_sample_character_limit: int = 8000`).
  - `_REFERENCE_LINE_MARKERS` tuple replaced by `ReferenceLineMarker` enum iteration.
  - Magic numbers in paragraph slicing (`3`, `2`, `3`, `2`) named as module-level constants
    (`_INTRODUCTION_PARAGRAPH_COUNT`, `_MIDDLE_PARAGRAPH_COUNT`, `_CONCLUSION_PARAGRAPH_LIMIT`,
    `_FALLBACK_TAIL_PARAGRAPH_COUNT`).
- Zero `os`/`dotenv` imports — verify by reading the file's import block directly.
- **Satisfies**: Requirement "QualityTextSampler Owns Text-Sampling Logic" (all scenarios) and
  Requirement "Tunable Sampling Values Are Constructor Parameters, Not Domain Env Reads" (zero-
  environment-import scenario).
- **Parallel/Sequential**: Sequential after T-02 (needs `ReferenceLineMarker`) and T-07 (RED).

---

## Phase 2 — `QualityResponseParser` (new file, ported logic)

### T-09: RED — write `QualityResponseParser` tests (ported from old `test_quality_analyzer.py`)
- File: `src/domain/tests/quality/test_quality_response_parser.py` — new file, calls
  `QualityResponseParser().parse(response_text)` directly, asserting on the returned
  `ParsedResponseDTO.scores[dimension].score` / `.feedback` instead of a full `analyze()` result.
- Port these existing TestCase classes/methods, adapted to call `parser.parse(...)` directly:
  - `TestHeaderFormats.test_numbered_and_unnumbered_headers_both_parse_to_same_score`
  - `TestNarrativeScoreInference.test_score_inferred_from_narrative_when_explicit_score_absent`
  - `TestNarrativeScoreInference.test_excelente_keyword_infers_eight_point_five`
  - `TestNarrativeScoreInference.test_aceptable_keyword_infers_six_point_zero`
  - `TestNarrativeScoreInference.test_deficiente_keyword_infers_four_point_zero`
  - `TestNarrativeScoreInference.test_no_keyword_match_uses_neutral_default_score`
  - `TestFeedbackExtraction.test_feedback_shorter_than_ten_characters_becomes_neutral_default`
  - `TestFeedbackExtraction.test_feedback_longer_than_three_sentences_is_truncated`
  - `TestDimensionMapping.test_argumentacion_block_is_not_misclassified_as_claridad`
  - `TestDimensionMapping.test_one_missing_dimension_in_otherwise_valid_response_keeps_the_rest`
    (adapted to assert on `ParsedResponseDTO.matched_dimensions` directly instead of inferring
    from a full `analyze()` call — this directly exercises the "one missing dimension" parsing
    scenario without needing 2 LLM calls).
- **Satisfies**: Requirement "QualityResponseParser Owns Per-Dimension Response Parsing" (all 5
  scenarios) and Requirement "DimensionScoreDTO and ParsedResponseDTO Extracted as Domain DTOs"
  ("parse() returns ParsedResponseDTO" half).
- **Parallel/Sequential**: Sequential RED before T-10 GREEN. Can be written in parallel with
  Phase 0 and Phase 1's RED (T-07).

### T-10: GREEN — implement `QualityResponseParser`
- File: `src/domain/quality/quality_response_parser.py` — exact code from design.md, ported
  verbatim from the current `_parse_response`/`_extract_score`/`_infer_score_from_narrative`/
  `_extract_feedback`/`_map_block_to_dimension` in `quality_analyzer.py`.
- 3 regex patterns + `_UNSCORED_DIMENSION_*` constants + `_NARRATIVE_SCORE_KEYWORDS` remain named
  module-level constants in this file (not enums — per spec's explicit reasoning).
- `parse()` returns `ParsedResponseDTO`; internal per-block pairs are `DimensionScoreDTO`.
- **Satisfies**: Requirement "QualityResponseParser Owns Per-Dimension Response Parsing" (all
  scenarios).
- **Parallel/Sequential**: Sequential after T-04, T-06 (needs both DTOs) and T-09 (RED).

---

## Phase 3 — Prompt template files (no test; static text)

### T-11: Create prompt template files verbatim
- `src/infrastructure/resources/prompts/quality/clarity_coherence_prompt.txt`
- `src/infrastructure/resources/prompts/quality/argumentation_conclusions_prompt.txt`
- Exact Spanish text from design.md, copied verbatim from the current
  `_build_prompt_one`/`_build_prompt_two` f-string bodies in `quality_analyzer.py`, with
  `{text_sample}` retained as a literal `.format()`-style placeholder (identical syntax to the
  legacy f-string placeholder, so rendered output is byte-for-byte unchanged).
- Create `src/infrastructure/resources/__init__.py`,
  `src/infrastructure/resources/prompts/__init__.py`,
  `src/infrastructure/resources/prompts/quality/__init__.py` if the project's existing resource
  folders use `__init__.py` markers (confirm convention from an existing infra folder before
  adding — if no other infra subfolder uses `__init__.py` for non-importable resources, skip it).
- No test file for the `.txt` content itself (static text, not executable code) — verified
  indirectly by T-14's "rendered prompt preserves legacy wording" test, which loads these exact
  files' content as literal strings in the test (or an equivalent literal string matching them,
  per spec's note that domain tests use literal template strings, not file I/O).
- **Satisfies**: Requirement "Prompt Template Injection — Domain Stays File-I/O-Free" (file
  existence + exact content half — the rendering behavior is T-13/T-14).
- **Parallel/Sequential**: Parallel with Phase 0, Phase 1, Phase 2 (fully independent files).

---

## Phase 4 — `QualityAnalyzer` rewritten as thin orchestrator

### T-12: RED — write new `test_quality_analyzer.py` (overwritten, orchestration-only scope)
- File: `src/domain/tests/quality/test_quality_analyzer.py` — overwrite in place once Phase 5's
  redistribution is confirmed complete (do not delete prematurely; see T-17 for the actual
  deletion/overwrite step). For this task, draft the new content as a separate working file or
  directly replace content with the orchestration-only scenarios below — construction now injects
  real `QualityTextSampler()` / `QualityResponseParser()` instances (no fakes needed) plus a fake
  `LlmGeneratorPort` and 2 literal prompt template strings containing `{text_sample}`.
- Keep ONLY these orchestration-level scenarios (ported/adapted from the old file, scope-narrowed
  per design.md's "Test Doubles (updated)" section):
  - `test_generate_is_called_exactly_twice_per_analysis` (from old
    `TestPortCallCountAndOverallScore`).
  - `test_claridad_and_coherencia_always_come_from_call_one` (from old
    `TestDirectPerCallAssignment`).
  - `test_argumentacion_and_conclusiones_always_come_from_call_two` (from old
    `TestDirectPerCallAssignment`).
  - `test_both_dimensions_failing_to_parse_in_one_call_raises_quality_analysis_failed` (from old
    `TestFullPerCallParseFailure`).
  - `test_overall_score_is_mean_of_four_dimension_scores` (from old
    `TestPortCallCountAndOverallScore`).
  - `test_overall_score_of_seven_resolves_to_good_quality_level` (from old
    `TestPortCallCountAndOverallScore`).
  - `test_domain_service_has_zero_infrastructure_imports` (from old
    `TestPortCallCountAndOverallScore`).
- Add 2 new tests required by the updated constructor shape:
  - `test_rendered_prompt_preserves_legacy_wording_with_sample_interpolated` — construct
    `QualityAnalyzer` with a literal template string containing `{text_sample}` plus legacy
    Spanish wording fragments; assert the prompt passed to the fake port's `generate()` contains
    that wording verbatim with the built sample substituted.
  - `test_quality_analyzer_module_defines_exactly_one_class` — inspect `quality_analyzer.py`'s
    top-level class definitions (e.g. via `ast` or `inspect`), assert exactly one:
    `QualityAnalyzer`.
- Do NOT keep: `TestTextSampling`, `TestConclusionDetection`, `TestHeaderFormats`,
  `TestNarrativeScoreInference`, `TestFeedbackExtraction`, `TestDimensionMapping` — these moved to
  T-07 and T-09 respectively.
- **Satisfies**: Requirement "QualityAnalyzer Domain Service Is a Thin Orchestrator" (all 3
  scenarios), Requirement "Prompt Template Injection — Domain Stays File-I/O-Free" (rendering
  scenario + "two methods collapse into one" scenario, the latter verified by T-19's grep),
  Requirement "Direct Per-Call Dimension Assignment, No Cross-Call Heuristic" (both scenarios),
  Requirement "Full Per-Call Parse Failure Raises QualityAnalysisFailed" (full-failure scenario —
  partial-failure scenario already covered by T-09's parser-level test), Requirement "Overall
  Score and Quality Level Computation" (both scenarios).
- **Parallel/Sequential**: Sequential after T-08, T-10, T-11 (constructor needs
  `QualityTextSampler`, `QualityResponseParser`, and literal prompt text fragments to exist/be
  known). Sequential RED before T-13 GREEN.

### T-13: GREEN — rewrite `QualityAnalyzer` as thin orchestrator
- File: `src/domain/quality/quality_analyzer.py` — full rewrite per design.md's exact code.
- Constructor: `llm_generator: LlmGeneratorPort`, `text_sampler: QualityTextSampler`,
  `response_parser: QualityResponseParser`, `clarity_coherence_prompt_template: str`,
  `argumentation_conclusions_prompt_template: str`.
- `analyze()`: delegates sampling to `text_sampler.build_sample()`, renders both prompts via one
  private `_render_prompt(template, text_sample)` helper, calls `generate()` exactly twice,
  delegates parsing to `response_parser.parse()` per call, keeps
  `_ensure_call_produced_usable_content` validation, direct per-call dimension assignment,
  averages into `overall_score`, maps via `get_quality_level_from_score()`, returns
  `QualityResultDTO`.
- Exactly 1 class in the file (`QualityAnalyzer`) — `_DimensionScore`/`_ParsedResponse` removed
  entirely, all regex/sampling logic removed entirely (now lives in T-08/T-10's files).
- **Satisfies**: same requirements as T-12.
- **Parallel/Sequential**: Sequential after T-12.

---

## Phase 5 — Cleanup: delete superseded code and old test file

### T-14: Verify no remaining references to deleted private classes/methods
- Grep the full `src/` tree for `_DimensionScore`, `_ParsedResponse`, `_build_text_sample`,
  `_collect_conclusion_or_tail_paragraphs`, `_is_reference_like`, `_build_prompt_one`,
  `_build_prompt_two`, `_parse_response`, `_extract_score`, `_infer_score_from_narrative`,
  `_extract_feedback`, `_map_block_to_dimension` (the old private names — note
  `QualityResponseParser`'s new methods reuse some of these names internally, which is fine; this
  check targets `quality_analyzer.py` specifically).
- Confirm `quality_analyzer.py` contains none of these identifiers after T-13.
- **Satisfies**: prerequisite for Requirement "QualityAnalyzer Domain Service Is a Thin
  Orchestrator" — "exactly one class" scenario, confirming full removal, not just file rewrite.
- **Parallel/Sequential**: Sequential after T-13.

### T-15: Confirm `test_quality_analyzer.py` no longer references removed scenarios
- Read the post-T-12 `test_quality_analyzer.py` in full; confirm every test method maps to one of
  T-12's listed orchestration scenarios and nothing from the old sampling/parsing TestCase classes
  remains.
- **Satisfies**: prerequisite for T-17 (confirms safe to treat the old file's content as fully
  superseded before final deletion bookkeeping).
- **Parallel/Sequential**: Sequential after T-12 (can run before or after T-13/T-14 — independent
  check).

### T-16: Verify `QualityResponseParser`/`QualityTextSampler` test coverage parity
- Cross-check each TestCase/method ported into T-07 and T-09 against the original
  `test_quality_analyzer.py` (read from git history or the pre-rework working tree) — confirm
  every one of these 9 original TestCase classes has every method accounted for in the new
  location:
  `TestTextSampling`, `TestConclusionDetection`, `TestHeaderFormats`,
  `TestNarrativeScoreInference`, `TestFeedbackExtraction`, `TestDimensionMapping`,
  `TestDirectPerCallAssignment`, `TestFullPerCallParseFailure`,
  `TestPortCallCountAndOverallScore`.
- Produce an explicit checklist (in the PR description or a scratch note, not a new repo file) of
  old-method → new-file mapping so no scenario is silently dropped.
- **Satisfies**: "Risks / Open Items for Tasks Phase" item in design.md — explicit redistribution
  tracking requirement.
- **Parallel/Sequential**: Sequential after T-07 and T-09 are both written (their RED content must
  exist to cross-check against).

### T-17: Final cleanup — no orphaned old test file
- Since `test_quality_analyzer.py` is overwritten in place (T-12), not created alongside an old
  copy, there is no separate "old file" to delete — confirm via `git status`/`git diff` that the
  file's history shows a modification (not a stray duplicate old-content file left anywhere in
  the tree).
- **Satisfies**: housekeeping — confirms the strangler approach completed cleanly with no leftover
  duplicate test files.
- **Parallel/Sequential**: Sequential after T-15, T-16.

---

## Phase 6 — Cross-cutting verification (sequential, after all GREEN)

### T-18: Verify one-class-per-file on `quality_analyzer.py`
- Parse `quality_analyzer.py`'s AST (or visually inspect); assert exactly one top-level class:
  `QualityAnalyzer`.
- **Satisfies**: Requirement "QualityAnalyzer Domain Service Is a Thin Orchestrator" — "exactly
  one class" scenario (final confirmation after T-13/T-14).
- **Parallel/Sequential**: Parallel with T-19, T-20.

### T-19: Verify `_build_prompt_one`/`_build_prompt_two` no longer exist anywhere
- Grep `quality_analyzer.py` for `_build_prompt_one` and `_build_prompt_two`; assert zero matches.
- Confirm a single `_render_prompt` helper exists and is called exactly twice inside `analyze()`.
- **Satisfies**: Requirement "Prompt Template Injection — Domain Stays File-I/O-Free" — "two
  prompt-building methods collapse into one" scenario.
- **Parallel/Sequential**: Parallel with T-18, T-20.

### T-20: Verify zero file I/O / env imports in `src/domain/`
- Grep `quality_analyzer.py`, `quality_text_sampler.py`, `quality_response_parser.py` for `open(`,
  `import os`, `from os`, `dotenv`; assert zero matches across all 3 files.
- **Satisfies**: Requirement "Prompt Template Injection — Domain Stays File-I/O-Free" (zero file
  I/O in domain) and Requirement "Tunable Sampling Values Are Constructor Parameters, Not Domain
  Env Reads" (zero environment/dotenv imports scenario).
- **Parallel/Sequential**: Parallel with T-18, T-19.

### T-21: ruff check on all new/modified files
- Run `ruff check` against: `reference_line_marker.py`, `dimension_score_dto.py`,
  `parsed_response_dto.py`, `quality_text_sampler.py`, `quality_response_parser.py`,
  `quality_analyzer.py`, and all 5 new/modified test files.
- Fix any lint findings before proceeding to T-22.
- **Satisfies**: general code-quality gate, not requirement-specific.
- **Parallel/Sequential**: Sequential after T-18, T-19, T-20.

### T-22: Full regression suite run
- Run `python -m pytest src/`.
- Expect 0 regressions vs the 273-test PR-A baseline. Exact final count may shift slightly
  (a handful of tests split or consolidated during redistribution, e.g. T-07's 2 new
  constructor-parameter tests are net additions; some narrative-inference sub-cases may already
  have been counted as separate methods) — the hard requirement is zero net loss of distinct
  behavioral assertions, not an exact number match.
- **Satisfies**: overall rework acceptance gate — confirms the strangler refactor preserved all
  behavior while restructuring.
- **Parallel/Sequential**: Sequential, last task.

---

## Review Workload Forecast

- **Estimated changed/added lines**: ~480-540 lines, all confined to `src/domain/` and 2 static
  resource files — no infrastructure/application/wiring changes in this rework (those are
  untouched from PR-A's already-merged-in-spirit shape and remain PR-B scope). Breakdown:
  - New production files: `reference_line_marker.py` (~10), `dimension_score_dto.py` (~12),
    `parsed_response_dto.py` (~15), `quality_text_sampler.py` (~50), `quality_response_parser.py`
    (~75). Subtotal: ~162.
  - Rewritten production file: `quality_analyzer.py` shrinks from ~240 lines to ~75 lines (net
    **negative** diff for this file, but the diff itself — old lines removed + new lines added —
    is still substantial since nearly the entire file changes).
  - 2 new resource `.txt` files: ~20 lines combined (no Python, but still new files in the diff).
  - New test files: `test_reference_line_marker.py` (~15), `test_dimension_score_dto.py` (~20),
    `test_parsed_response_dto.py` (~25), `test_quality_text_sampler.py` (~90, ports 3 existing
    scenarios + 2 new constructor-parameter tests), `test_quality_response_parser.py` (~140, ports
    10 existing scenarios). Subtotal: ~290.
  - Rewritten test file: `test_quality_analyzer.py` shrinks from ~382 lines to ~90-110 lines
    (7 ported + 2 new orchestration-only scenarios).
  - **Total estimated diff**: ~480-540 lines changed/added (sum of removals in
    `quality_analyzer.py`/`test_quality_analyzer.py` plus additions across 9 new files), noting
    this is a refactor where a large fraction of the diff is deletion, not net new surface area.

- **400-line budget risk**: **Medium**. Raw added+removed line count likely exceeds 400, but the
  net new behavioral surface is small (no new requirements, no new external dependencies, no new
  infrastructure) — this is structurally a redistribution of already-reviewed, already-tested
  logic across more files, not new logic requiring the same scrutiny depth as PR-A's original
  review. The reviewer's job is "confirm the move preserved behavior," not "evaluate new business
  rules."

- **Decision needed before apply**: **Yes, but with a recommendation**. Per `delivery_strategy:
  ask-on-risk`, the orchestrator MUST stop and ask the user whether to:
  1. **Land as additional commits on the SAME PR-A branch/PR #13** (RECOMMENDED) — since PR #13
     has not merged yet, these commits are pre-merge revisions to work the user has not yet
     approved into `main`. Splitting this rework into its own separate PR would create a strange
     review sequence (reviewer approves PR #13's monolith, then immediately reviews a PR that
     deletes most of what was just approved). Stacking these commits onto PR #13 directly lets the
     reviewer see one coherent final diff against `main` reflecting the user's actual intended
     design, with the rework commits readable as the natural "iterate before merge" history.
  2. Open a separate PR-A2 that depends on PR #13 merging first, then immediately supersedes its
     internal structure (NOT recommended for the same reason as above — adds review overhead for
     no benefit since PR #13 hasn't merged).
  3. Proceed with a `size:exception` label if commit-stacking onto PR #13 is not feasible for
     branch-protection or CI reasons.
  Do not silently choose — surface this explicitly, but note the strong structural preference for
  option 1 given PR #13's unmerged state.
