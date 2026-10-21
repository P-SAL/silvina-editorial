# Delta Spec — check-grammar

**Change**: resolve-magic-values-debt

## MODIFIED Requirements

### Requirement: LanguageToolAdapter

`LanguageToolAdapter(GrammarCheckPort)` MUST:
- Import `language_tool_python` at module level.
- Accept a required `max_replacements: int` in its constructor (no default — the sole default lives in the wiring's env-var read) and store it.
- Store `self._tool: language_tool_python.LanguageTool | None = None`; initialize `LanguageTool('es')` inside `check()` on first call only.
- Sample first 20 paragraphs, truncated to 5000 chars total before passing to LanguageTool.
- Filter out matches where `rule_issue_type == 'misspelling'`.
- Return at most the first 10 errors as `list[GrammarErrorDTO]`.
- Limit the replacements (suggestions) in each `GrammarErrorDTO` to at most `max_replacements`.
- Catch exceptions raised during `LanguageTool` initialization or `check()` and re-raise as `GrammarCheckUnavailable` (manual try/except; `@generic_error_handler` is not used here, since it is scoped to the use-case boundary, not adapters).

(Previously: The replacement suggestions limit was hardcoded to 3.)

#### Scenario: Lazy init — no Java on import
- GIVEN the `language_tool_adapter` module is imported
- WHEN no `check()` call has been made
- THEN `self._tool` is `None` on any adapter instance

#### Scenario: Misspelling results are filtered
- GIVEN LanguageTool returns 3 grammar errors and 2 misspellings
- WHEN `check(paragraphs)` processes the results
- THEN exactly 3 `GrammarErrorDTO` instances are returned

#### Scenario: Output capped at 10 errors
- GIVEN LanguageTool returns 12 grammar errors (no misspellings)
- WHEN `check(paragraphs)` processes the results
- THEN exactly 10 `GrammarErrorDTO` instances are returned

#### Scenario: Backend failure propagates GrammarCheckUnavailable
- GIVEN `LanguageTool('es')` raises any exception during init or check
- WHEN `check(paragraphs)` is called
- THEN `GrammarCheckUnavailable` is raised (via the manual try/except in `check()`/`_initialize_tool_if_needed`)

#### Scenario: Replacements limit is respected
- GIVEN a `LanguageToolAdapter` initialized with `max_replacements` = 2
- AND LanguageTool returns an error with 5 suggestions
- WHEN `check(paragraphs)` processes the results
- THEN the returned `GrammarErrorDTO` contains at most 2 replacements

---

### Requirement: AnalyzeDocumentUseCaseWiring Wires the Grammar Port Directly

`AnalyzeDocumentUseCaseWiring` MUST expose a private `_get_grammar_check_port() -> GrammarCheckPort` returning a `LanguageToolAdapter` constructed with `max_replacements` loaded from the environment variable `GRAMMAR_MAX_REPLACEMENTS` (defaulting to 3). No business logic in the wiring class.

(Previously: `_get_grammar_check_port()` did not load any grammar settings or pass `max_replacements`.)

#### Scenario: Wiring produces correctly typed instance
- GIVEN `AnalyzeDocumentUseCaseWiring()` is instantiated
- WHEN `create_use_case()` is called
- THEN the resulting `AnalyzeDocumentUseCase`'s `_grammar_check_port` is a `LanguageToolAdapter`
- AND the adapter is configured with the value from environment variable `GRAMMAR_MAX_REPLACEMENTS` if defined
