# Technical Debt Registry — silvina-editorial

Consolidated inventory of technical debt accumulated during the hexagonal migration (Slices 0-15). Sourced from Engram decisions and `openspec/archive/*` artifacts, cross-checked against current code on `refactor/hexagonal-migration` (2026-07-02).

Each item lists status as verified against the current codebase, not just the state at the time it was logged.
## Confirmed still present

### 1. Magic values not audited across the remaining `src/` infrastructure adapters (partially resolved)
The `ArticleSize` enum (Spanish members) and the magic-number thresholds in `ArticleSizeClassifier`, `QualityLevelResolver`, and `PublicationVerdictEvaluator` were resolved — see "Resolved" section below. The following magic literals identified during that audit were explicitly deferred (out of scope for that change) and remain unaddressed:
- `StructureValidator._extract_present_sections()` — hardcoded `100` char header limit.
- `DocxCitationAdapter._collect_multi_author()` — hardcoded `100` max author name length.
- `DocxReferenceAdapter` — raw regex flags `2 | 16` instead of `re.IGNORECASE | re.DOTALL`.
- `LanguageToolAdapter._map_to_dto()` — hardcoded `3` max replacements.
- `DocxReportAdapter` — `250` words/page divisor, `5` list slice size, `150` truncation size, `3` replacements limit, `6`x`2` table dimensions, `80` separator repeat count.
- **Source**: Engram #630; `openspec/changes/archive/2026-07-04-resolve-spanish-and-magic-debt/exploration.md`.

### 2. Spanish domain-vocabulary enums pending final rename pass (partially resolved)
`ArticleType`'s Spanish member KEYS were renamed to English — see "Resolved" section below. The following remain deliberately untouched:
- `QualityDimension` — keys already English; `.value`s stay Spanish (match literal LLM/document text), rename of values still deferred to the final pass together with parsing logic.
- `SectionType` — keys and values both stay Spanish/bilingual on purpose: members like `RESUMEN`/`ABSTRACT`, `INTRODUCCION`/`INTRODUCTION` are intentional parallel-language pairs, not translation debt.
- **Source**: Engram #613; `openspec/changes/archive/2026-07-04-resolve-domain-vocabulary-enums-debt/`.

### 3. Accumulated dead-code registry (by design — not to be cleaned per-slice)
Per convention (Engram #605), dead code found while migrating a legacy module is documented, not removed, and batched for a future dedicated cleanup pass instead of being fixed slice-by-slice:
- Slice 4 (validate-citations): `extract_all_citations()` in `citation_matcher.py` — no call sites, confirmed dead. `business_logic/article_analyzer.py` (`ArticleAnalyzer`) — whole module never wired to `main.py`/`gradio_app.py`, sole caller of `generate_report()`.
- Slice 5 (analyze-quality): `self.client = ollama.Client(...)` in `QualityAnalyzer.__init__` — built but never used. `article_type` param of `analyze_quality(document_content, article_type)` — never read in the method body (kept intentionally, not removed). `analyze_document_quality()` convenience function at end of file — calls the instance method with 3 args against a 2-arg signature; broken/unreachable.
- **Source**: Engram #605 (topic `migration/dead-code-registry`).

### 5. README.md documents the deleted legacy structure
`README.md` (lines 18, 63-66, 215-236) still describes the old 4-layer legacy root layout (`domain/`, `data_access/`, `business_logic/`, `presentation/`, `apa_validator.py`, `eumic_verifier.py`) that Slice 16 (`cleanup-legacy-packages`) deleted from the repo. The "Project Structure" section shows a directory tree that no longer exists.
- Update to describe the `src/` hexagonal layout (`src/domain/`, `src/application/`, `src/infrastructure/`) and remove references to the deleted legacy packages/files.
- **Source**: judgment-day review of `cleanup-legacy-packages` PR2 (2026-07-05) — flagged by Judge B, verified real via grep; deferred by user decision to a later pass.

### 6. `@generic_error_handler` applied to adapters instead of use cases only
Convention: `@generic_error_handler` should only decorate application-layer use case methods, never infrastructure adapters — error handling is a use-case responsibility, not an adapter one. The following adapters currently violate this:
- `DocxTextAdapter.read_paragraphs()` (`src/infrastructure/adapters/document/docx_text_adapter.py`)
- `DocxCitationAdapter.extract_citations()` (`src/infrastructure/adapters/document/docx_citation_adapter.py`)
- `DocxReferenceAdapter.extract_references()` (`src/infrastructure/adapters/document/docx_reference_adapter.py`)
- `DocxEumicAdapter.inspect()` (`src/infrastructure/adapters/document/docx_eumic_adapter.py`)
- `OllamaGeneratorAdapter.generate()` (`src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py`)
- Pre-existing debt, not introduced by `refactor_analyze_document_wiring` — none of these 5 files were touched by that change.
- **Source**: Engram `pattern/generic-error-handler-scope`; discovered during judgment-day review of `refactor_analyze_document_wiring` (2026-07-04).

## Resolved (was tracked, no longer applies)

### Legacy modules (`domain/`, `data_access/`, `business_logic/`, `presentation/`) not yet deleted
Slice 16 (final cleanup: delete legacy top-level packages) was postponed until the full hexagonal migration (Slices 0-15) was confirmed working, then executed as its own SDD change.
- **Verified fixed**: 2026-07-05, openspec change `cleanup-legacy-packages` (PR1 commit `54967c8` + PR2 commit `60ba349`) — deleted `domain/`, `data_access/`, `business_logic/` (incl. `vocab/`), `presentation/`, `apa_validator.py`, `eumic_verifier.py`, `config.py`, `main_legacy.py`, `tests/legacy/`; adapted `tests/smoke/`/`tests/e2e/` to test `src/` directly; cleaned dead `ruff.toml` exclude entries. Full suite green: 589 passed, 3 skipped, 0 failed. Both PRs passed judgment-day (PR1 approved after 1 fix round, PR2 approved with 1 deferred follow-up — see item 5 above).
- **Source**: Engram #735 (original), #802-#810 (proposal, design, tasks, apply/verify PR1+PR2, judgment-day PR1); `openspec/changes/cleanup-legacy-packages/`.

### Explicit keyword arguments not audited across the full codebase
Audited and refactored the codebase to use explicit keyword arguments in all function/method calls to prevent bugs and enforce clean interfaces.
- **Verified fixed**: 2026-07-03, openspec change `resolve-explicit-keyword-arguments` — verified cleanly by `sdd-verify`.
- **Source**: `openspec/changes/archive/2026-07-03-resolve-explicit-keyword-arguments/`.

### Method-ordering convention not audited across existing classes
Reordered methods in existing classes to follow the convention of public methods before private, grouped alphabetically (excluding dunder methods).
- **Verified fixed**: 2026-07-03, openspec change `resolve-method-ordering` — verified cleanly by `sdd-verify`.
- **Source**: `openspec/changes/archive/2026-07-03-resolve-method-ordering/`.

### Slice 5 OOP/encapsulation violations
Fixed encapsulation and OOP patterns in Slice 5 code, including `QualityLevelResolver` extraction, pattern encapsulation in `QualityTextSampler`, and cleaner test helpers.
- **Verified fixed**: 2026-07-03, openspec change `resolve-slice-5-technical-debt` — verified cleanly by `sdd-verify`.
- **Source**: `openspec/changes/archive/2026-07-03-resolve-slice-5-technical-debt/`.

### Bare module imports in `main.py` and `tests/test_main_cli_args.py`
Prohibited full-module imports like `import os` and replaced them with specific name imports (e.g. `from os import getenv`).
- **Verified fixed**: 2026-07-03, openspec change `resolve-bare-imports` — verified cleanly by `sdd-verify`.
- **Source**: `openspec/changes/archive/2026-07-03-resolve-bare-imports/`.

### `tests/smoke/test_validate_structure_parity.py` broken import
Was importing `DocumentContent` from `src.domain.dtos.document_content_dto`, but the real class is `DocumentContentDTO`. Logged as tech debt in `openspec/archive/2026-07-01-refactor-slice-14-cli/tasks.md`.
- **Verified fixed**: 2026-07-02 commit `890221f` ("fix: resolve issues found during Slice 15 manual QA") — file now imports `DocumentContentDTO` correctly (`tests/smoke/test_validate_structure_parity.py:26`).

### `main()` didn't propagate `save_word_report()` failure to exit code
`main.py` called `silvina.save_word_report(...)` without checking the returned bool, so a failed Word export still printed "ANÁLISIS COMPLETADO" and exited 0.
- **Verified fixed**: 2026-07-02, openspec change `resolve-cli-and-fake-debt` — `main.py` now saves the JSON report first (no data loss), then checks the bool and does `sys.exit(1)` with an explicit error message if the Word report failed. Covered by `tests/test_main_cli_args.py::TestMainExitCodes::test_exits_1_when_save_word_report_fails`. Verified independently by `sdd-verify`.
- **Source**: `openspec/archive/2026-07-01-refactor-slice-14-cli/tasks.md`; Engram #723; `openspec/changes/resolve-cli-and-fake-debt/`.

### `FakeLlmGeneratorPort` naming violated hexagonal terminology
`src/domain/tests/quality/fake_llm_generator_port.py` named its class `FakeLlmGeneratorPort` without inheriting from `LlmGeneratorPort` (pure duck-typing) — a fake implementation of a Port is an Adapter, not a Port.
- **Verified fixed**: 2026-07-02, openspec change `resolve-cli-and-fake-debt` — file renamed to `fake_llm_generator_adapter.py`, class renamed to `FakeLlmGeneratorAdapter(LlmGeneratorPort)`, signature aligned to `generate(self, prompt: str, options: dict | None = None)`. `test_quality_analyzer.py` updated. Zero residual references confirmed by `sdd-verify`.
- **Source**: Engram #631; `openspec/changes/resolve-cli-and-fake-debt/`.

### Bare `pytest` collection fails — `domain` package name collision
Running `pytest -q` from the repo root threw `ModuleNotFoundError: No module named 'domain.tests'` across dozens of files under `src/domain/tests/`. Both the legacy top-level `domain/` package and `src/domain/` shared the name `domain`, causing pytest's default import mode to resolve the wrong one (104 collection errors on baseline).
- **Verified fixed**: 2026-07-02, openspec change `resolve-pytest-domain-collision` — configured `pytest.ini` with `addopts = --import-mode=importlib` and `pythonpath = .`, added empty `src/__init__.py` to prevent top-level namespace collision. Results: 635 passed, 3 skipped, 0 collection errors (confirmed via `.venv/Scripts/pytest.exe -q`). Legacy `domain/` and `src/domain/` remain untouched as required.
- **Source**: Engram #752-#756 (proposal, design, tasks, apply-progress, verify-report); `openspec/changes/resolve-pytest-domain-collision/`.

### `ArticleSize` enum Spanish members and magic-number thresholds in classification/quality/recommendation
`ArticleSize` enum members were in Spanish (`LARGO`, `CORTO`, `NO_DEFINIDO`, `FUERA_RANGO`); `ArticleSizeClassifier`, `QualityLevelResolver`, and `PublicationVerdictEvaluator` had hardcoded magic-number thresholds.
- **Verified fixed**: 2026-07-04, openspec change `resolve-spanish-and-magic-debt` — enum members renamed to `LONG`/`SHORT`/`UNDEFINED`/`OUT_OF_RANGE` (Spanish `.value`s preserved for report/downstream compatibility); the three classes now accept keyword-only injected thresholds (defaults unchanged), loaded from `.env` via their wirings/config. No behavioral changes. `.venv/Scripts/pytest.exe -q` → 641 passed, 3 skipped (up from 635 baseline). Verified independently by `sdd-verify` (0 CRITICAL).
- **Source**: Engram #630, #777-#783 (proposal, spec, design, tasks, apply-progress, verify-report, archive-report); `openspec/changes/archive/2026-07-04-resolve-spanish-and-magic-debt/`.

### `ArticleType` enum Spanish member keys
`ArticleType` used Spanish identifiers (`CIENTIFICO`, `DIVULGACION`) for its enum members.
- **Verified fixed**: 2026-07-04, openspec change `resolve-domain-vocabulary-enums-debt` — keys renamed to `SCIENTIFIC`/`POPULAR_SCIENCE` (Spanish `.value`s preserved for report/downstream compatibility). Scope restricted to `src/`: the legacy `domain/enums.py` mirror and `business_logic/*` keep their original Spanish keys (independent enum, not imported by `src/`). `SectionType` was explicitly left out of scope — its Spanish members are intentional bilingual pairs, not debt. No behavioral changes. `.venv/Scripts/pytest.exe -q` → 641 passed, 3 skipped. Verified independently by `sdd-verify` (0 CRITICAL).
- **Source**: Engram #613, #786-#793 (proposal, spec, design, tasks, verify-report, archive-report); `openspec/changes/archive/2026-07-04-resolve-domain-vocabulary-enums-debt/`.

## Notes on scope

- Inline code search (`TODO`/`FIXME`/`HACK`/`XXX`) across `src/` and the repo root found no real markers — all tracked debt lives in Engram decisions and `openspec/archive/*` artifacts, not code comments.
- This document does not create a new SDD change (no proposal/spec/design/tasks cycle) — it's a point-in-time consolidated snapshot for planning a future dedicated cleanup pass, per the project's existing convention of batching debt instead of fixing it per-slice.
