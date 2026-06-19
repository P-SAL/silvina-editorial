# Tasks: analyze-quality (Slice 5)

Phased TDD structure matching Slices 2-4, extended with a SCAFFOLD phase for the
two brand-new top-level folders (`src/domain/ports/`,
`src/infrastructure/adapters/`) introduced for the first time by this slice.

Pre-existing and reused as-is (confirmed by reading the repo, not re-created):
- `src/domain/exceptions/quality_errors.py` — `QualityAnalysisFailed` already defined.
- `src/domain/exceptions/language_model_errors.py` — `LanguageModelUnavailable` already defined.
- `src/domain/dtos/quality_result_dto.py` — `QualityResultDTO` already defined, unchanged.
- `src/domain/enums/quality_level.py` — `QualityLevel` enum already defined; this slice only
  appends `get_quality_level_from_score()` to it.

---

## Phase 0 — SCAFFOLD (sequential, first-time folder creation)

### T-01: Create `src/domain/ports/` package [x]
- Create `src/domain/ports/__init__.py` (empty, matches existing package convention).
- **Satisfies**: prerequisite for Requirement "LlmGeneratorPort Contract".
- **Parallel/Sequential**: Sequential — must exist before T-02.

### T-02: Create `src/infrastructure/adapters/llm_generator/` package
- Create `src/infrastructure/adapters/__init__.py` and
  `src/infrastructure/adapters/llm_generator/__init__.py`.
- **Satisfies**: prerequisite for Requirement "OllamaGeneratorAdapter Implements the Port".
- **Parallel/Sequential**: Sequential — must exist before T-05/T-06. Can run in parallel with T-01
  (different folders).

### T-03: Create `src/domain/quality/` entity folder [x]
- Create `src/domain/quality/__init__.py` (mirrors `domain/structure/`, `domain/citation/`
  pattern from prior slices).
- **Satisfies**: prerequisite for Requirement "QualityAnalyzer Domain Service Depends Only on
  the Port".
- **Parallel/Sequential**: Parallel with T-01, T-02.

### T-04: Create `src/domain/tests/quality/` and confirm `src/domain/tests/enums/` exist [x]
- Create `src/domain/tests/quality/__init__.py`.
- Confirm `src/domain/tests/enums/` exists (used in T-12); create if missing.
- **Parallel/Sequential**: Parallel with T-01–T-03.

---

## Phase 1 — RED/GREEN: `LlmGeneratorPort` (domain, no test required)

### T-05: Write `LlmGeneratorPort` Protocol [x]
- File: `src/domain/ports/llm_generator_port.py`.
- `Protocol` with exactly one method: `generate(self, prompt: str) -> str`.
- No test file — this is a pure structural `Protocol` declaration with zero logic. Per skill
  conventions, a test would only assert "method exists with this signature," which provides no
  behavioral coverage; the contract is verified transitively by `OllamaGeneratorAdapter`'s tests
  (T-07) and `QualityAnalyzer`'s fake-double tests (T-13–T-21). **Explicitly noting this
  deviation from the "every file gets a test" default**, justified because there is no branch,
  no computation, no constructor logic to exercise.
- **Satisfies**: Requirement "LlmGeneratorPort Contract" (both scenarios verified by inspection/
  static signature, no runtime test needed).
- **Parallel/Sequential**: Sequential after T-01.

---

## Phase 2 — RED/GREEN: `OllamaGeneratorAdapter`

### T-06: RED — write `OllamaGeneratorAdapter` tests (mocked `ollama`)
- File: `src/infrastructure/tests/test_ollama_generator_adapter.py`.
- Mock `ollama.generate` at module level inside `ollama_generator_adapter.py` (no real Ollama
  server contacted).
- Test cases:
  - `test_generate_returns_stripped_response_text` — mocked `ollama.generate` returns
    `{'response': '  some text  '}`; assert adapter returns `"some text"`.
  - `test_generate_raises_language_model_unavailable_on_backend_failure` — mocked
    `ollama.generate` raises a generic exception (e.g. `ConnectionError`); assert
    `LanguageModelUnavailable` is raised.
- **Satisfies**: Requirement "OllamaGeneratorAdapter Implements the Port" — both scenarios.
- **Parallel/Sequential**: Sequential after T-05, T-02.

### T-07: GREEN — implement `OllamaGeneratorAdapter`
- File: `src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py`.
- Implements `LlmGeneratorPort`. Module-level `ollama.generate()` call (no `Client` field, per
  ADR-2 — out of scope, dead code not carried over).
- `generate()` decorated with `@generic_error_handler`; explicit `try/except Exception as exc:
  raise LanguageModelUnavailable() from exc` wraps the `ollama.generate()` call itself.
- Returns `response.get('response', '').strip()`.
- This file is the ONLY new file importing `ollama` — verify with a grep/inspection step before
  marking done (supports the spec's "sole Ollama import site" scenario).
- **Satisfies**: Requirement "OllamaGeneratorAdapter Implements the Port" (all 3 scenarios,
  including "Adapter is the sole Ollama import site" — verified once all slice files exist, see
  T-22).
- **Parallel/Sequential**: Sequential after T-06 (TDD RED before GREEN).

---

## Phase 3 — RED/GREEN: `QualityDimension` enum

### T-08: RED — write `QualityDimension` enum test [x]
- File: `src/domain/tests/enums/test_quality_dimension.py`.
- `test_enum_has_exactly_four_members` — assert `len(QualityDimension) == 4`.
- `test_enum_members_are_claridad_coherencia_argumentacion_conclusiones` — assert each named
  member exists with expected string value (`"claridad"`, `"coherencia"`, `"argumentacion"`,
  `"conclusiones"`).
- `test_analysis_dimension_enum_is_unchanged` — import existing `AnalysisDimension`, assert
  `QualityDimension` does not subclass or alias it (e.g. `assert not issubclass(QualityDimension,
  AnalysisDimension)` style check, or simply assert they remain two distinct classes with no
  shared base beyond `Enum`).
- **Satisfies**: Requirement "QualityDimension Enum" — both scenarios.
- **Parallel/Sequential**: Parallel with T-05/T-06/T-07 (independent file). Sequential RED before
  T-09 GREEN.

### T-09: GREEN — implement `QualityDimension` enum [x]
- File: `src/domain/enums/quality_dimension.py`.
- 4 members: `CLARIDAD = "claridad"`, `COHERENCIA = "coherencia"`,
  `ARGUMENTACION = "argumentacion"`, `CONCLUSIONES = "conclusiones"`.
- Does not touch `src/domain/enums/analysis_dimension.py` (or wherever `AnalysisDimension`
  lives) at all.
- **Satisfies**: Requirement "QualityDimension Enum".
- **Parallel/Sequential**: Sequential after T-08.

---

## Phase 4 — RED/GREEN: `get_quality_level_from_score`

### T-10: RED — write threshold tests for `get_quality_level_from_score` [x]
- File: `src/domain/tests/enums/test_quality_level.py` (new file — confirmed no existing test
  file for `quality_level.py` in current tree; create fresh, do not assume one exists).
- Test cases covering all 5 boundary thresholds:
  - `test_score_of_nine_point_zero_returns_excellent` (`>= 9.0`)
  - `test_score_just_below_nine_returns_good` (e.g. `8.9`)
  - `test_score_of_seven_point_zero_returns_good` (`>= 7.0`)
  - `test_score_just_below_seven_returns_acceptable` (e.g. `6.9`)
  - `test_score_of_five_point_zero_returns_acceptable` (`>= 5.0`)
  - `test_score_just_below_five_returns_needs_improvement` (e.g. `4.9`)
  - `test_score_of_three_point_zero_returns_needs_improvement` (`>= 3.0`)
  - `test_score_just_below_three_returns_poor` (e.g. `2.9`)
- **Satisfies**: Requirement "Overall Score and Quality Level Computation" — quality-level
  boundary scenario.
- **Parallel/Sequential**: Parallel with T-08/T-09 (independent file). Sequential RED before
  T-11 GREEN.

### T-11: GREEN — append `get_quality_level_from_score()` to `quality_level.py` [x]
- File: `src/domain/enums/quality_level.py` (append function, do not touch existing
  `QualityLevel` enum body).
- Module-level function per design ADR (companion function to the enum it lives next to).
- **Satisfies**: Requirement "Overall Score and Quality Level Computation".
- **Parallel/Sequential**: Sequential after T-10.

---

## Phase 5 — RED/GREEN: `QualityAnalyzer` domain service (largest phase)

All `QualityAnalyzer` tests use a fake `LlmGeneratorPort` test double (structural duck-typing,
no real Ollama, no real adapter). Define the fake once, reuse across test cases.

### T-12: Define fake `LlmGeneratorPort` test double for domain tests [x]
- File: `src/domain/tests/quality/test_quality_analyzer.py` (fake class defined inline at top of
  test file, or as a small local helper class — not a production test double, scoped to domain
  tests only; this differs from the `infrastructure/tests/test_doubles/` wiring-level double in
  T-21, which is for wiring tests).
- Fake supports scripting per-call return values (ordered list or call-count-indexed responses)
  and records call count for the "called exactly twice" assertion.
- **Satisfies**: prerequisite for all T-13–T-20 tests.
- **Parallel/Sequential**: Sequential before T-13. Can run in parallel with Phases 1-4.

### T-13: RED/GREEN — text sampling: short document fallback [x]
- Test: `test_short_document_uses_full_text_fallback_instead_of_sample` — document whose
  strategically sampled text totals fewer than 400 words; assert the prompt built from it
  contains the full joined paragraph text, not just the sample slice.
- **Satisfies**: Requirement "Text Sampling and Prompt Construction Preserved Verbatim" — short
  document fallback scenario.

### T-14: RED/GREEN — text sampling: long document uses strategic sample [x]
- Test: `test_long_document_uses_strategic_sample_not_full_text` — document whose sample totals
  400+ words; assert the prompt is built from title+intro+middle+conclusion sample, not the full
  document text (e.g. assert a paragraph excluded from the sample window does not appear in the
  prompt).
- **Satisfies**: same requirement — long document scenario.

### T-15: RED/GREEN — conclusion detection excludes reference-like lines [x]
- Test: `test_conclusion_paragraphs_exclude_reference_like_lines` — paragraphs after a detected
  "conclusi..." paragraph, some containing `http`/`doi.org`/`https`/`ISBN` in their first 80
  characters; assert those lines are excluded from the collected conclusion paragraphs.
- **Satisfies**: same requirement — reference-exclusion scenario.

### T-16: RED/GREEN — parsing: numbered and unnumbered headers [x]
- Test: `test_numbered_and_unnumbered_headers_both_parse_to_same_score` — one response using
  `**1. Claridad...` and another using `**Claridad...`; assert both yield score `8.0` with
  consistent feedback extraction.
- **Satisfies**: Requirement "Per-Dimension Response Parsing Preserved Verbatim" — header format
  scenario.

### T-17: RED/GREEN — parsing: narrative score inference [x]
- Test: `test_score_inferred_from_narrative_when_explicit_score_absent` — block containing "bueno
  y adecuado" with no `[Puntuación: X/10]` pattern; assert inferred score is `7.5`.
- Also cover the other 3 narrative buckets (`excelente`→`8.5`, `aceptable`→`6.0`,
  `deficiente`→`4.0`) and the neutral-default no-keyword-match case in the same or sibling test
  methods.
- **Satisfies**: same requirement — narrative inference scenario.

### T-18: RED/GREEN — parsing: feedback shorter than 10 chars becomes neutral default [x]
- Test: `test_feedback_shorter_than_ten_characters_becomes_neutral_default`.
- **Satisfies**: same requirement — short-feedback scenario.

### T-19: RED/GREEN — parsing: feedback truncated to 3 sentences [x]
- Test: `test_feedback_longer_than_three_sentences_is_truncated`.
- **Satisfies**: same requirement — truncation scenario.

### T-20: RED/GREEN — parsing: argumentacion vs claridad disambiguation + partial failure [x]
- Test: `test_argumentacion_block_is_not_misclassified_as_claridad` — block containing both
  `argumentaci` and `argumento` in first 200 chars; assert mapped to `ARGUMENTACION`.
- Test: `test_one_missing_dimension_in_otherwise_valid_response_keeps_the_rest` — Call 1 with
  valid Claridad block but no Coherencia header; assert Claridad uses parsed value, Coherencia
  falls back to neutral default, no exception raised.
- **Satisfies**: Requirement "Per-Dimension Response Parsing Preserved Verbatim" (dimension
  mapping order scenario) AND Requirement "Full Per-Call Parse Failure Raises
  QualityAnalysisFailed" (partial-failure-does-not-raise scenario).

### T-20b: RED/GREEN — direct per-call assignment, no cross-call heuristic [x]
- Test: `test_claridad_and_coherencia_always_come_from_call_one` — Call 1 has valid
  Claridad/Coherencia; Call 2's text also happens to match a Claridad-like header; assert final
  Claridad/Coherencia come from Call 1 only.
- Test: `test_argumentacion_and_conclusiones_always_come_from_call_two`.
- **Satisfies**: Requirement "Direct Per-Call Dimension Assignment, No Cross-Call Heuristic" —
  both scenarios.

### T-20c: RED/GREEN — full per-call parse failure raises `QualityAnalysisFailed` [x]
- Test: `test_both_dimensions_failing_to_parse_in_one_call_raises_quality_analysis_failed` — Call
  1's response contains neither a Claridad nor Coherencia header anywhere; assert
  `QualityAnalysisFailed` is raised.
- **Satisfies**: Requirement "Full Per-Call Parse Failure Raises QualityAnalysisFailed" — full
  failure scenario.

### T-20d: RED/GREEN — port called exactly twice; overall score and quality level [x]
- Test: `test_generate_is_called_exactly_twice_per_analysis` — fake port records call count;
  assert exactly 2 after one `analyze()` call.
- Test: `test_overall_score_is_mean_of_four_dimension_scores` — final scores `8.0, 6.0, 7.0, 9.0`
  → assert `overall_score == 7.5`.
- Test: `test_overall_score_of_seven_resolves_to_good_quality_level`.
- Test: `test_domain_service_has_zero_infrastructure_imports` — inspect
  `quality_analyzer.py` source/imports for absence of `src.infrastructure` or `ollama`.
- **Satisfies**: Requirement "QualityAnalyzer Domain Service Depends Only on the Port" (both
  scenarios) AND Requirement "Overall Score and Quality Level Computation" (both scenarios).

### T-21: GREEN — implement `QualityAnalyzer` [x]
- File: `src/domain/quality/quality_analyzer.py`.
- Implements per design.md verbatim: `_ParsedResponse`/`_DimensionScore` private dataclasses,
  `_build_text_sample`, `_collect_conclusion_or_tail_paragraphs`, `_is_reference_like`,
  `_build_prompt_one`/`_build_prompt_two` (exact Spanish text from legacy), `_parse_response`,
  `_extract_score`, `_infer_score_from_narrative`, `_extract_feedback`,
  `_map_block_to_dimension`, `_ensure_call_produced_usable_content`, public `analyze()`.
- Run all of T-13 through T-20d against this single implementation — this is one GREEN step
  covering the whole class, since the methods are mutually dependent (cannot meaningfully
  implement `_parse_response` without `_extract_score`/`_extract_feedback`/
  `_map_block_to_dimension` all present).
- **Satisfies**: all `QualityAnalyzer`-related requirements listed in T-13–T-20d.
- **Parallel/Sequential**: Sequential after T-12 and after all RED tests in T-13–T-20d are
  written (classic single-GREEN-for-multiple-RED-tests pattern, consistent with how Slices 2-4
  implemented multi-method domain services).

---

## Phase 6 — RED/GREEN: `AnalyzeQualityUseCase`

### T-22: RED — write `AnalyzeQualityUseCase` tests
- File: `src/application/tests/test_analyze_quality_use_case.py`.
- `test_execute_returns_quality_analyzer_result_unchanged` — fake/stub `QualityAnalyzer` (or
  reuse the fake `LlmGeneratorPort` + real `QualityAnalyzer`) confirms the use case's returned
  `QualityResultDTO` matches what `QualityAnalyzer.analyze` produces for the same input.
- `test_execute_accepts_different_article_type_values_without_error` — two calls with same
  `document_content`, different `article_type` values; both succeed.
- **Satisfies**: Requirement "AnalyzeQualityUseCase Thin Pass-Through" — both scenarios.
- **Parallel/Sequential**: Sequential after T-21 (depends on real `QualityAnalyzer` or a fake of
  it).

### T-23: GREEN — implement `AnalyzeQualityUseCase`
- File: `src/application/analyze_quality_use_case.py`.
- Thin pass-through per design.md; keeps unused `article_type` parameter in the signature.
- **Satisfies**: Requirement "AnalyzeQualityUseCase Thin Pass-Through".
- **Parallel/Sequential**: Sequential after T-22.

---

## Phase 7 — RED/GREEN: `AnalyzeQualityUseCaseWiring`

### T-24: RED — write `AnalyzeQualityUseCaseWiring` tests + test double
- File: `src/infrastructure/tests/test_analyze_quality_use_case_wiring.py`.
- Test double file: `src/infrastructure/tests/test_doubles/analyze_quality_use_case_wiring_for_test.py`
  (`AnalyzeQualityUseCaseWiringForTest` overriding `_get_llm_generator_port()` with a fake,
  following the Slices 2-4 `WiringForTest` pattern).
- `test_create_use_case_returns_ready_to_use_analyze_quality_use_case` — calling
  `create_use_case()` on the **production** wiring returns an `AnalyzeQualityUseCase` instance
  backed by a real `OllamaGeneratorAdapter` (assert via type inspection of the injected adapter,
  not a live Ollama call).
- `test_llm_generator_port_accessor_returns_port_type_annotation` — inspect
  `_get_llm_generator_port`'s return type annotation; assert it is `LlmGeneratorPort`, not
  `OllamaGeneratorAdapter`.
- **Satisfies**: Requirement "AnalyzeQualityUseCaseWiring Assembles Domain Service and Adapter" —
  both scenarios.
- **Parallel/Sequential**: Sequential after T-23 and T-07 (needs both `AnalyzeQualityUseCase` and
  `OllamaGeneratorAdapter` to exist).

### T-25: GREEN — implement `AnalyzeQualityUseCaseWiring`
- File: `src/infrastructure/wirings/analyze_quality_use_case_wiring.py`.
- Single public method `create_use_case()` (confirmed exact name match with
  `ValidateStructureWiring`, `ValidateApaWiring` — *correction*: read `match_citations_use_case_
  wiring.py` directly and confirmed `create_use_case()` is the consistent name across all
  existing wirings — `MatchCitationsUseCaseWiring`, `ValidateApaWiring`. Use `create_use_case()`,
  not `get_<use_case>()`).
- `_get_quality_analyzer()` returns `QualityAnalyzer`; `_get_llm_generator_port()` returns
  `LlmGeneratorPort` (port type annotation, concrete `OllamaGeneratorAdapter()` instance
  returned inside).
- **Satisfies**: Requirement "AnalyzeQualityUseCaseWiring Assembles Domain Service and Adapter".
- **Parallel/Sequential**: Sequential after T-24.

---

## Phase 8 — Cross-cutting verification (sequential, after all GREEN)

### T-26: Verify no `print()` calls in any new file
- Grep all 8 new production files (`llm_generator_port.py`, `ollama_generator_adapter.py`,
  `quality_dimension.py`, `quality_level.py` diff only, `quality_analyzer.py`,
  `analyze_quality_use_case.py`, `analyze_quality_use_case_wiring.py`) for `print(`.
- **Satisfies**: Requirement "No print() Statements in Migrated Code".
- **Parallel/Sequential**: Sequential, after T-25 (all files must exist).

### T-27: Verify `ollama` import isolation
- Grep all new files for `import ollama` / `from ollama`; assert only
  `ollama_generator_adapter.py` matches.
- **Satisfies**: Requirement "OllamaGeneratorAdapter Implements the Port" — "sole Ollama import
  site" scenario (closing the loop opened in T-07).
- **Parallel/Sequential**: Parallel with T-26.

### T-28: Full regression suite run
- Run `python -m pytest src/`.
- Expect 242 (baseline) + all new tests from T-06, T-08, T-10, T-13 through T-20d, T-22, T-24 —
  0 regressions, 0 failures.
- **Satisfies**: overall slice acceptance — no requirement-specific mapping, this is the
  integration gate for the whole slice.
- **Parallel/Sequential**: Sequential, last task.

---

## Review Workload Forecast

- **Estimated changed/added lines**: ~520-600 lines across 8 new production files + 7 new test
  files + 2 new `__init__.py` packages + 1 modified file (`quality_level.py`, ~12-line append).
  Breakdown:
  - Production: `llm_generator_port.py` (~10), `ollama_generator_adapter.py` (~35),
    `quality_dimension.py` (~12), `quality_level.py` append (~12), `quality_analyzer.py`
    (~200, largest file in the slice — port + adapter + enum + domain service + use case +
    wiring combined), `analyze_quality_use_case.py` (~10),
    `analyze_quality_use_case_wiring.py` (~15). Subtotal: ~294.
  - Tests: `test_ollama_generator_adapter.py` (~40), `test_quality_dimension.py` (~25),
    `test_quality_level.py` (~40), `test_quality_analyzer.py` (~180-220, largest test file —
    covers sampling, parsing, dimension assignment, failure modes, call-count), `test_analyze_
    quality_use_case.py` (~30), `test_analyze_quality_use_case_wiring.py` (~30),
    `analyze_quality_use_case_wiring_for_test.py` (~10). Subtotal: ~355-395.
  - Total estimate: **~650-690 lines**, across **17 new files + 1 modified file**.

- **400-line budget risk**: **High**. This slice's estimate (~650-690 lines) exceeds the ~400-line
  PR budget by roughly 60-70%, driven by:
  1. Two brand-new top-level folders requiring scaffolding overhead not present in Slices 2-4.
  2. `quality_analyzer.py` and its test file are both significantly larger than any single
     component in Slices 2-4 (legacy file is 247 lines with dense regex parsing logic; porting it
     verbatim plus full test coverage of every parsing branch is inherently large).
  3. This is the first slice with a port+adapter pair, doubling the infrastructure surface
     (port file + adapter file + adapter test + wiring test double) compared to pure-domain
     slices.

- **Chained PRs recommended**: **Yes**. Suggested split (2 PRs, each independently mergeable and
  testable):
  - **PR-A (domain + port)**: T-01, T-03, T-04, T-05 (port), T-08/T-09 (`QualityDimension`),
    T-10/T-11 (`get_quality_level_from_score`), T-12 through T-21 (`QualityAnalyzer` + all its
    tests, using the fake port double — no real adapter needed). Estimated ~470-510 lines.
  - **PR-B (adapter + application + wiring)**: T-02, T-06/T-07 (`OllamaGeneratorAdapter`), T-22/
    T-23 (`AnalyzeQualityUseCase`), T-24/T-25 (`AnalyzeQualityUseCaseWiring`), T-26 through T-28
    (final verification + full regression). Estimated ~180-200 lines. Depends on PR-A merging
    first (`QualityAnalyzer` and `LlmGeneratorPort` must exist).
  - This mirrors the PR-A/PR-B split already used for Slice 2 (`migration/slice2-validate-
    structure-domain` then `-app`), so it is consistent with established precedent in this
    migration, not a new pattern.

- **Decision needed before apply**: **Yes** — per `delivery_strategy: ask-on-risk`, the
  orchestrator MUST stop and ask the user whether to:
  (a) split into the PR-A/PR-B chain above (and if so, confirm `chain_strategy`:
  `stacked-to-main` vs `feature-branch-chain`), or
  (b) proceed as a single PR with an explicit `size:exception` label.
  Do not silently choose either path.
