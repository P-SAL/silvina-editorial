# Proposal: classify-article (Slice 6)

## Intent

`business_logic/structure_analyzer.py` (`StructureAnalyzer`, ~35 lines) and
`business_logic/article_classifier.py` (`ArticleClassifier`, ~280+ lines) together decide a
document's `ArticleType` (CIENTIFICO/DIVULGACION/OPINION) using a 19-case deterministic rule
table fed by 6 signals — most of them pure/local (IMRyD section-keyword detection, reference
count/recency, methodological vocabulary), one of them LLM-backed (a single combined Ollama
call extracting 3 yes/no answers). This is the **second slice requiring real I/O**, and the
first opportunity to validate that `LlmGeneratorPort` (established in `analyze-quality`,
Slice 5) generalizes to a second, different LLM consumer rather than being an
quality-analysis-specific abstraction in disguise. The migration must preserve the legacy
classifier's exact decision behavior (rule table, confidence values, IMRyD override) while
moving it into the hexagonal `domain/application/infrastructure` split, and must resolve a
naming collision the legacy code itself does not have to deal with: the already-migrated
`StructureValidator` (a different, unrelated domain service from `validate-structure`,
Slice 4) makes the literal name `StructureAnalyzer` ambiguous in `src/domain/`.

## Scope

### In Scope

- **Rename `StructureAnalyzer` to avoid collision with `StructureValidator`.** The migrated
  class must NOT be named `StructureAnalyzer` — it solves a different problem
  (deterministic IMRyD section-keyword *presence* detection, producing a boolean signal dict
  consumed only by classification) than `src/domain/structure/structure_validator.py`'s
  `StructureValidator` (per-`ArticleType` required-section completeness, producing
  `(present, missing)` lists for `validate-structure`'s own use case). The two having
  "Structure"-prefixed names side by side in `src/domain/` would actively mislead future
  readers into assuming overlap or a refactor opportunity that does not exist. Exact name
  (e.g. `ImrydSignalDetector` or equivalent) is finalized during design; the constraint that
  it must NOT be `StructureAnalyzer` is fixed now. This is an explicit, scoped exception to
  the "don't rename pre-existing legacy identifiers" convention — justified by an unrelated
  same-prefix collision the convention's authors did not anticipate, not a general license to
  rename other legacy names.
- `src/domain/classification/` (or equivalent package, finalized in design) housing:
  - The renamed IMRyD signal-detection service — single method
    `detect(document: DocumentContentDTO) -> dict[str, bool]` (or equivalent), ported verbatim
    from `StructureAnalyzer.analyze()`, including the bilingual `IMRYD_KEYWORDS` table and the
    `imryd_complete` flag's exact semantics (requires intro/methods/results/discussion; does
    NOT require conclusion despite detecting it).
  - `ArticleClassifier` (or equivalent name) — the orchestrating domain service, constructor-
    injected with `LlmGeneratorPort`, the renamed signal detector, a text sampler, and a
    response parser, mirroring `QualityAnalyzer`'s constructor-injection pattern from Slice 5.
  - A dedicated text sampler for classification (legacy `_build_text_sample()`: first 3500 +
    last 2500 chars, skips bibliography via short-paragraph marker detection) — ported as its
    own injectable unit, distinct from `QualityTextSampler` (different sampling algorithm, no
    behavior change to either).
  - A response parser for the combined S4/S5/S6 LLM call (regex-extracts 3 yes/no answers
    from one free-text response).
  - The 19-case `_apply_rule` decision table, ported **verbatim** — no business-rule
    reinterpretation, no accuracy improvements. Per-branch Spanish-prose `reasoning` strings
    are preserved exactly as legacy produces them.
- **New `ClassificationConfidence` enum** (`float`-mixin Enum, e.g.
  `class ClassificationConfidence(float, Enum)`), replacing the 5 distinct confidence literals
  (0.95/0.90/0.86/0.85/0.83) scattered as inline magic numbers across the legacy rule table.
  Each member is usable directly as a float (comparisons, arithmetic, DTO serialization) while
  being self-documenting (e.g. `EXACT_MATCH = 0.95`). This is a brand-new enum introduced
  during migration, so per the established convention its members get **English** names —
  unlike `ArticleType`, which is a pre-existing legacy identifier and keeps its Spanish member
  names (`CIENTIFICO`/`DIVULGACION`/`OPINION`) unchanged.
- **Additive `options` parameter on `LlmGeneratorPort.generate()`.** Extend the signature from
  `generate(self, prompt: str) -> str` to
  `generate(self, prompt: str, options: dict | None = None) -> str`, defaulting to `None` so
  `analyze-quality`'s existing call site (`generate(prompt)`) is completely unaffected — no
  changes required to `QualityAnalyzer` or its tests. `OllamaGeneratorAdapter.generate()`
  forwards `options` straight through to `ollama.generate(options=options)` (the underlying
  `ollama` library already accepts this kwarg natively for tuning such as
  `temperature`/`num_predict`). This is a deliberate, scoped reopening of a port shipped in
  Slice 5 — justified because `classify-article` is the first second consumer to prove the
  port's actual reuse shape, and the legacy classifier's `temperature=0.1` tuning was a
  deliberate low-variance choice for its yes/no signal extraction that should not be silently
  dropped during migration.
- `ArticleClassifier` domain service receives `temperature` and `num_predict` as constructor
  parameters (same pattern as `QualityTextSampler`'s tunables) — **no defaults in the domain
  service or adapter constructors**. Defaults live only in classify-article's own use-case
  wiring, sourced from `.env` (e.g. `ARTICLE_CLASSIFIER_TEMPERATURE`,
  `ARTICLE_CLASSIFIER_NUM_PREDICT`), exactly mirroring the `analyze-quality` precedent for
  tunable values. `ArticleClassifier` passes these as
  `options={"temperature": ..., "num_predict": ...}` on its `generate()` call.
- Migrate the orphaned `classify_article_size(char_count) -> ArticleSize` helper function,
  confirmed present in legacy `domain/enums.py` but NOT yet carried into
  `src/domain/enums/article_size.py` (only the bare `ArticleSize` enum was migrated in an
  earlier slice). Added to that same file, mirroring how `quality_level.py` already carries
  `get_quality_level_from_score`.
- `src/application/classify_article_use_case.py` — `ClassifyArticleUseCase.execute(document_content: DocumentContentDTO) -> ClassificationResultDTO`. Thin pass-through to the domain
  service, same shape as `AnalyzeQualityUseCase`.
- `src/infrastructure/wirings/classify_article_use_case_wiring.py` —
  `ClassifyArticleUseCaseWiring`, instance-based `_get_*` pattern, reusing
  `OllamaGeneratorAdapter` (already exists from Slice 5 — no new adapter needed) and reading
  the new `ARTICLE_CLASSIFIER_TEMPERATURE`/`ARTICLE_CLASSIFIER_NUM_PREDICT` env vars.
- Externalize the long Spanish S4/S5/S6 prompt text to
  `src/infrastructure/resources/prompts/classification/`, following the `PROMPTS_DIR` package
  pattern established in `analyze-quality`.
- Reuse `ClassificationResultDTO`, `DocumentContentDTO`, `ReferenceDTO`, `ArticleType`,
  `ArticleSize` as-is — all already exist in `src/domain/`, no new DTOs needed beyond the one
  new enum (`ClassificationConfidence`) described above.
- Domain tests with a fake `LlmGeneratorPort` test double (no real Ollama calls), covering the
  IMRyD override path and the 19-case rule table with per-signal-combination granularity
  (finer than `validate-structure`'s per-`ArticleType` test-file split, given classification has
  19 distinct cases vs. structure-validation's 4). Adapter tests for the extended `generate()`
  signature confirm `options` forwards correctly and that omitting it preserves
  `analyze-quality`'s existing behavior unchanged.
- `tests/smoke/test_classify_article_parity.py`, following the exact pattern already
  established by `tests/smoke/test_validate_structure_parity.py` (Slice 4): runs the legacy
  `ArticleClassifier` and the new `ClassifyArticleUseCaseWiring`-built use case against the
  same real `.docx` sample documents in `docs/sample-documents/`, asserting identical
  `(article_type, confidence)` output. The Ollama call on BOTH sides is patched with the same
  fixed canned response (e.g. `"S4: SI\nS5: SI\nS6: SI"`) via `unittest.mock.patch` — this
  preserves the unique value of the original `__main__` smoke test (real document parsing
  exercised end-to-end) while requiring no live external service, per explicit user
  constraint: no test may depend on a running Ollama instance.

### Out of Scope

- `business_logic/structure_validator.py` / `src/domain/structure/structure_validator.py` —
  already migrated in a separate slice (`validate-structure`). No code changes; cross-checked
  for naming/logic overlap with the renamed IMRyD signal detector, confirmed none beyond the
  shared naming-collision concern this proposal resolves by renaming.
- `business_logic/article_analyzer.py` — the final top-level orchestrator that will call
  `ClassifyArticleUseCase`, `AnalyzeQualityUseCase`, and the validation use cases together.
  Confirmed already drifted/stale relative to current APIs (calls a `classifier.classify(dict)`
  method that does not exist on current `ArticleClassifier`). Deferred to a future slice that
  depends on this one.
- Any accuracy improvements, rule-table changes, or new classification signals. **Exact
  behavioral parity with the legacy classifier is the goal of this slice.** The 19-case rule
  table migrates verbatim. Future precision improvements are explicitly deferred to a separate
  future change.
- Fixing any production bugs or edge cases in the classification logic — there are none known.
  This is a pure architectural migration of working legacy behavior, not a bugfix exercise.
- Deleting `business_logic/structure_analyzer.py` / `business_logic/article_classifier.py` —
  coexistence maintained until the caller-switchover slice.
- Wiring `ClassifyArticleUseCase` into `main.py` — deferred to the caller-switchover slice
  (likely alongside or after `article_analyzer.py`'s own migration).
- Deduplicating the IMRyD signal detector's keyword-matching logic against
  `StructureValidator`'s independent section-alias matching — both heuristically reimplement
  similar paragraph-scanning logic today; a shared abstraction is a real but separate
  refactor opportunity, not in scope here (would touch `validate-structure`'s already-merged
  code).
- The legacy classifier's own unused `self.client = ollama.Client(...)`-style direct
  dependency and any `print()`-based progress messages — dropped per the no-`print()`-in-
  domain-code rule, consistent with `analyze-quality`'s precedent, not replaced with logging
  in this slice.
- Module-level `analyze_structure()` / `classify_document()` convenience functions — legacy
  scaffolding, not carried over, matching the `analyze-quality` precedent. Confirmed via
  repo-wide grep that neither is called from `main.py`, `gradio_app.py`, or any other module.
- `structure_analyzer.py`'s `__main__` smoke-test block — purely synthetic (inline paragraphs,
  no real document I/O) and asserts nothing (just prints). No unique value beyond what
  `ImrydSignalDetector`'s domain unit tests already cover with real assertions; not carried
  over.
- `article_classifier.py`'s `__main__` smoke-test block in its original form (raw `assert`
  statements gated behind `if __name__ == "__main__":`) is NOT carried over verbatim — but see
  the new in-scope item below, since unlike `structure_analyzer.py`'s block, this one has
  unique value (real `.docx` document parsing) worth preserving in a proper test.

## Capabilities

### New Capabilities

- `classify-article`: deterministic + LLM-backed article classification producing a
  `ClassificationResultDTO` (article type, size, confidence, reasoning), exposed via
  `ClassifyArticleUseCase`. Second capability in the migration with an external dependency,
  reusing `LlmGeneratorPort` (extended additively) rather than introducing a new port.

### Modified Capabilities

- `LlmGeneratorPort` / `OllamaGeneratorAdapter` (introduced in `analyze-quality`, Slice 5):
  `generate()` gains an additive, optional `options: dict | None = None` parameter. Existing
  `analyze-quality` callers and tests are unaffected — this is a backward-compatible signature
  extension, not a breaking change.

## Approach — Why This Slice Extends Rather Than Duplicates Slice 5's Port

`analyze-quality` deliberately named its port `LlmGeneratorPort` (generic, capability-based)
rather than something quality-specific, explicitly anticipating that a future classification
slice would also need text generation. This slice is that anticipated second consumer, and it
validates — rather than contradicts — that naming choice: both `QualityAnalyzer` and
`ArticleClassifier` need "send a prompt, get text back," confirming the port's shape was
correctly minimal. The one place the original port falls short is generation-option tuning:
the legacy classifier deliberately calls Ollama with `temperature=0.1, num_predict=300` for
low-variance yes/no signal extraction, a tuning choice not exercised by `analyze-quality`'s
call (which uses Ollama's defaults). Rather than forking a parallel
`ConfigurableLlmGeneratorPort` (rejected — would fragment the one-port precedent for no
structural reason) or silently dropping the tuning (rejected — changes the legacy's calibrated
behavior with no compensating benefit), the port gains one additive, optional parameter.

1. **Port extension** (`src/domain/ports/llm_generator_port.py`) — `generate()` gains
   `options: dict | None = None`. Zero impact on existing implementers/callers due to the
   default.
2. **Adapter extension** (`src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py`)
   — forwards `options` to `ollama.generate(options=options)` unmodified; `ollama` already
   supports this kwarg natively, no adapter-side translation logic needed.
3. **Renamed signal-detection service** — pure, deterministic, zero dependency on the LLM
   port; same testability profile as `StructureValidator`.
4. **`ArticleClassifier` domain service** — receives the port, the signal detector, sampler,
   and parser via constructor injection; also receives `temperature`/`num_predict` as plain
   constructor parameters (no defaults), passed through as `options={...}` on its one
   `generate()` call. Contains the 19-case rule table and all signal-computation logic
   verbatim, with zero knowledge of `ollama`.
5. **Use case + wiring** — same shape as Slice 5; wiring now also owns the
   `ARTICLE_CLASSIFIER_TEMPERATURE`/`ARTICLE_CLASSIFIER_NUM_PREDICT` env-var defaults,
   following the exact "defaults only in wiring" precedent already used for
   `QualityTextSampler`'s tunables.

## Affected Areas

| Area | Impact | Description |
|------|--------|--------------|
| `src/domain/ports/llm_generator_port.py` | Modified | Additive `options: dict \| None = None` parameter on `generate()` |
| `src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py` | Modified | Forwards `options` to `ollama.generate()` |
| `src/domain/classification/` (naming finalized in design) | New | Renamed IMRyD signal detector, `ArticleClassifier`, text sampler, response parser |
| `src/domain/enums/classification_confidence.py` (naming finalized in design) | New | `ClassificationConfidence(float, Enum)` — 5 members, English names |
| `src/domain/enums/article_size.py` | Modified | Adds the orphaned `classify_article_size()` helper function |
| `src/application/classify_article_use_case.py` | New | `ClassifyArticleUseCase` |
| `src/infrastructure/wirings/classify_article_use_case_wiring.py` | New | Wiring; reuses `OllamaGeneratorAdapter`, owns new env-var defaults |
| `src/infrastructure/resources/prompts/classification/` | New | Externalized S4/S5/S6 prompt, `PROMPTS_DIR` pattern |
| Domain/adapter tests (paths finalized in design) | New | Fake-port domain tests, real adapter `options`-forwarding tests |
| `business_logic/structure_analyzer.py` | Unchanged | Legacy stays alive during coexistence |
| `business_logic/article_classifier.py` | Unchanged | Legacy stays alive during coexistence |
| `src/domain/structure/structure_validator.py` | Unchanged | Confirmed no overlap; cross-checked only |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Reopening a port already shipped/merged in Slice 5 could regress `analyze-quality` | Low | Change is purely additive (`options` defaults to `None`); existing call sites and tests require zero modification; covered by adapter tests asserting omission preserves prior behavior |
| 19-case rule table is the largest, most complex logic in the slice; any subtle transcription error changes classification outcomes silently | Med | Port verbatim, no reinterpretation; cover with per-signal-combination tests (finer granularity than `validate-structure`'s per-type split, given 19 distinct cases) |
| Renaming `StructureAnalyzer` is an explicit, scoped exception to the no-rename-legacy-identifiers convention | Low | Justified narrowly by the unrelated name collision with `StructureValidator`; documented here so it is not mistaken for a precedent to rename other legacy identifiers |
| Dropping vs. preserving `temperature=0.1`/`num_predict=300` tuning affects yes/no signal-extraction determinism | Low (resolved) | User decision: preserve via additive port option, sourced from `.env` with wiring-level defaults — no domain/adapter defaults |
| `classify_article_size()` migration is easy to forget (orphaned, not co-located with `ArticleSize` enum today) | Low | Explicitly called out in scope; single function, low complexity |

## Rollback Plan

The port/adapter change is additive (new optional parameter, default `None`) — no rollback
needed for existing `analyze-quality` behavior even if classify-article's new files are
reverted. All other new files (domain package, enum, use case, wiring, prompts, tests) are
additive. Legacy `business_logic/structure_analyzer.py` and `business_logic/article_classifier.py`
are untouched. To roll back: delete the new domain package, enum, use case, wiring, prompt
resources, and test files; revert the additive `options` parameter on the port/adapter if
desired (safe either way since it is unused by any other caller until this slice). `main.py`
continues importing from `business_logic/`. No migration state to undo.

## Dependencies

- Slice 5 (`analyze-quality`) — establishes `LlmGeneratorPort`/`OllamaGeneratorAdapter`, the
  wiring `.env`-defaults-only-in-wiring pattern, and the `PROMPTS_DIR` externalized-prompt
  pattern, all reused/extended here.
- Slice 4 (`validate-structure`) — establishes `src/domain/structure/`, cross-checked for
  naming collision (resolved via rename) and logic overlap (none found, no dedup in scope).
- Slice 1 (`domain-exceptions`) — existing exception hierarchy available if classify-article
  needs to raise on LLM failure or unparseable response; exact exception(s) finalized in
  design (likely reusing `LanguageModelUnavailable` and/or a new
  `ClassificationFailed`-style domain exception).
- `ClassificationResultDTO`, `DocumentContentDTO`, `ReferenceDTO`, `ArticleType`, `ArticleSize`
  — already exist, reused as-is.
- `generic_error_handler` decorator — already exists, already applied to
  `OllamaGeneratorAdapter`; no change needed for the additive parameter.

## Success Criteria

- [ ] The migrated IMRyD signal-detection class is NOT named `StructureAnalyzer` (collision
      avoided with `StructureValidator`); exact name finalized in design
- [ ] `LlmGeneratorPort.generate()` signature is
      `generate(self, prompt: str, options: dict | None = None) -> str`; existing
      `analyze-quality` call sites and tests pass unmodified
- [ ] `OllamaGeneratorAdapter.generate()` forwards `options` to `ollama.generate(options=options)`
      verbatim, with no adapter-side defaults
- [ ] `ArticleClassifier`'s constructor requires `temperature`/`num_predict` with no defaults;
      defaults exist only in `classify_article_use_case_wiring.py`, sourced from
      `ARTICLE_CLASSIFIER_TEMPERATURE`/`ARTICLE_CLASSIFIER_NUM_PREDICT`
- [ ] New `ClassificationConfidence` enum has exactly 5 members with English names
      (e.g. `EXACT_MATCH = 0.95`), mixes in `float`, and is used everywhere the legacy rule
      table hardcoded 0.95/0.90/0.86/0.85/0.83
- [ ] `ArticleType` member names are unchanged (`CIENTIFICO`/`DIVULGACION`/`OPINION`) — no
      renaming applied to this pre-existing legacy identifier
- [ ] The 19-case rule table produces identical `(article_type, confidence, reasoning)` output
      to the legacy implementation for equivalent signal inputs — verified via tests covering
      each case
- [ ] `classify_article_size()` exists in `src/domain/enums/article_size.py` alongside
      `ArticleSize`
- [ ] No `print()` calls anywhere in the new domain/application code
- [ ] Legacy `business_logic/structure_analyzer.py` and `business_logic/article_classifier.py`
      are unmodified; `main.py` still imports from `business_logic/`

## Open Questions

None blocking — all naming, port-shape, and tuning-parameter decisions needed to start design
were resolved by explicit user decision during this proposal round (see the 5 constraints
encoded throughout the Scope and Approach sections above). Remaining items are pure
implementation/naming details correctly deferred to design:

1. **Exact class/package names** — final names for the renamed IMRyD signal detector,
   `ArticleClassifier`'s package location (`src/domain/classification/` vs.
   `src/domain/article_classification/` or similar), and the `ClassificationConfidence`
   enum's file path. Constraint already fixed: the signal detector must not be named
   `StructureAnalyzer`; the confidence enum must have English member names.
2. **Exception(s) raised on LLM/parsing failure** — whether classify-article reuses
   `LanguageModelUnavailable` as-is for adapter-level failures and introduces a new
   domain-level exception (mirroring `QualityAnalysisFailed`) for unparseable S4/S5/S6
   responses, or handles this differently. Design-level detail, not a scope question.
