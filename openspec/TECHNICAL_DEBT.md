# Technical Debt Registry — silvina-editorial

Consolidated inventory of technical debt accumulated during the hexagonal migration (Slices 0-15). Sourced from Engram decisions and `openspec/archive/*` artifacts, cross-checked against current code on `refactor/hexagonal-migration` (2026-07-02).

Each item lists status as verified against the current codebase, not just the state at the time it was logged.

## Confirmed still present

### 1. Bare `pytest` collection fails — `domain` package name collision
Running `pytest -q` from the repo root (no path scoping) throws `ModuleNotFoundError: No module named 'domain.tests'` across dozens of files under `src/domain/tests/`. Both the legacy top-level `domain/` package and `src/domain/` share the name `domain`, so pytest's default import mode resolves the wrong one.

- **Verified**: 2026-07-02, `pytest -q --collect-only` → 104 collection errors, 287 tests collected.
- **Mitigation in place**: always run scoped (`pytest src/ -q`, `pytest tests/ -q`), never the bare command.
- **Real fix pending**: `--import-mode=importlib`, or rename one of the two `domain` packages.
- **Source**: `openspec/archive/2026-07-01-refactor-slice-14-cli/tasks.md`.

### 2. Explicit keyword arguments not audited across the full codebase
Convention: every method call must use explicit keyword arguments, even for a single parameter. Applied going forward from Slice 7 (extract-citations) onward, but never audited retroactively across earlier slices.
- **Scope**: `src/application/`, `src/infrastructure/wirings/`, `src/infrastructure/adapters/`.
- **Source**: Engram #673 (topic `tech-debt/explicit-keyword-arguments`).

### 3. Method-ordering convention not audited across existing classes
Convention (defined during classify-article PR-1 review): public methods before private, no interleaving, alphabetical within each group, dunders/`__init__` exempt. Never applied retroactively to classes written before the rule existed (Slice 0, 5, 6).
- **Candidates**: `src/domain/quality/quality_analyzer.py`, `quality_text_sampler.py`, `quality_response_parser.py`; `src/domain/classification/article_classification_response_parser.py`, `article_classification_text_sampler.py`, `imryd_signal_detector.py`, `article_size_classifier.py`; `src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py`; `src/infrastructure/wirings/*.py`.
- **Source**: Engram #629.

### 4. Slice 5 OOP/encapsulation violations (deferred, not yet fixed)
Flagged during classify-article PR-1 review; explicitly left untouched to avoid scope creep into already-merged Slice 5 code:
- `src/domain/enums/quality_level.py` — `get_quality_level_from_score()` should become `QualityLevelResolver.resolve()`.
- `src/domain/quality/quality_text_sampler.py` — module-level `_CONCLUSION_HEADER_PATTERN` should move inside `QualityTextSampler` as a class attribute.
- `src/domain/tests/quality/test_quality_text_sampler.py` — module-level `build_document_content()` helper should become a private `TestCase` method.
- `src/application/tests/fake_llm_generator_adapter.py` — `FakeLlmGeneratorAdapterForTest.generate()` is missing the `options: dict | None = None` param added to `LlmGeneratorPort` in Slice 6 (ADR-4); latent signature drift, not currently exercised by any test passing `options=`.
- **Source**: Engram #627.

### 5. Non-English identifiers and magic values not audited across `src/`
Two pending audit items across the whole `src/` tree (not a single slice):
1. Any variable/constant/class/method name still in Spanish beyond the already-known cases (see item 6 below, which is the separate, already-scoped exception for domain-vocabulary enums).
2. Hardcoded/magic literals (numbers, strings) that should be named constants.
- **Source**: Engram #630.

### 6. Spanish domain-vocabulary enums pending final rename pass
Enums like `QualityDimension`, `ArticleType`, `SectionType` keep Spanish `.value`s (they match literal text from the LLM/documents) by deliberate decision — renaming is deferred to a dedicated final pass at the end of the whole migration, together with their parsing logic, not per-slice.
- **Source**: Engram #613.

### 7. Accumulated dead-code registry (by design — not to be cleaned per-slice)
Per convention (Engram #605), dead code found while migrating a legacy module is documented, not removed, and batched for a future dedicated cleanup pass instead of being fixed slice-by-slice:
- Slice 4 (validate-citations): `extract_all_citations()` in `citation_matcher.py` — no call sites, confirmed dead. `business_logic/article_analyzer.py` (`ArticleAnalyzer`) — whole module never wired to `main.py`/`gradio_app.py`, sole caller of `generate_report()`.
- Slice 5 (analyze-quality): `self.client = ollama.Client(...)` in `QualityAnalyzer.__init__` — built but never used. `article_type` param of `analyze_quality(document_content, article_type)` — never read in the method body (kept intentionally, not removed). `analyze_document_quality()` convenience function at end of file — calls the instance method with 3 args against a 2-arg signature; broken/unreachable.
- **Source**: Engram #605 (topic `migration/dead-code-registry`).

### 8. Legacy modules (`domain/`, `data_access/`, `business_logic/`, `presentation/`) not yet deleted
Slice 16 (final cleanup: delete legacy top-level packages) was explicitly postponed until the full hexagonal migration (Slices 0-15) is confirmed working in real use.
- This is why item 1 (the `domain` package collision) still exists — both the legacy and the new `src/domain` package coexist on purpose for now.
- **Source**: Engram #735.

### 9. Bare module imports in `main.py` and `tests/test_main_cli_args.py`
Convention (Engram #707, established during Slice 12 / export-report review): full-module imports like `import os` are prohibited — must import specific names instead (e.g. `from os.path import dirname, join, exists`). `main.py` and `tests/test_main_cli_args.py` predate or were never audited against this rule.
- **Verified**: 2026-07-02, `grep "^import "`:
  - `main.py:10-14,23` — `import argparse`, `import re`, `import sys`, `import os`, `import traceback`, `import json`.
  - `tests/test_main_cli_args.py:5-8` — `import io`, `import os`, `import sys`, `import unittest`.
- **Found by**: Judge A during the `resolve-cli-and-fake-debt` Judgment Day review (Round 2) — flagged as CRITICAL but ruled out of scope for that change since the diff didn't touch these lines; logged here instead.
- **Source**: Engram #707.

## Resolved (was tracked, no longer applies)

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

## Notes on scope

- Inline code search (`TODO`/`FIXME`/`HACK`/`XXX`) across `src/` and the repo root found no real markers — all tracked debt lives in Engram decisions and `openspec/archive/*` artifacts, not code comments.
- This document does not create a new SDD change (no proposal/spec/design/tasks cycle) — it's a point-in-time consolidated snapshot for planning a future dedicated cleanup pass, per the project's existing convention of batching debt instead of fixing it per-slice.
