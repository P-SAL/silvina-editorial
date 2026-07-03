# Tasks: Resolve Bare Imports

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 50-70 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | ask-on-risk |
| Chain strategy | stacked-to-main |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: stacked-to-main
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Refactor imports in main and test entrypoints | PR 1 | Base branch; contains code refactoring and test updates |

## Phase 1: Foundation / Infrastructure

- [x] 1.1 Run existing test suite using pytest to establish baseline pass state.

## Phase 2: Core Implementation

- [x] 2.1 Refactor imports in `main.py` to import specific names from `argparse`, `json`, `os.path`, `re`, `sys`, and `traceback`.
- [x] 2.2 Update bare module references in `main.py` to use imported specific names.
- [x] 2.3 Refactor imports in `tests/test_main_cli_args.py` to import specific names from `io`, `os.path`, `sys`, and `unittest`.
- [x] 2.4 Update path insertion and test class inheritance in `tests/test_main_cli_args.py` to use imported names.
- [x] 2.5 Lift local imports of `main` and `_build_argument_parser` to the top of `tests/test_main_cli_args.py` (after path bootstrapping).
- [x] 2.6 Refactor mock patching from `patch.object(sys, "argv", ...)` to `patch("sys.argv", ...)` in `tests/test_main_cli_args.py`.

## Phase 3: Testing / Verification

- [x] 3.1 Run tests using `.venv\Scripts\pytest tests/test_main_cli_args.py` to verify CLI args and entry point behavior.
- [x] 3.2 Run the full test suite using `.venv\Scripts\pytest` to verify no regression across other components.

## Phase 4: Cleanup / Documentation

- [x] 4.1 Run ruff linter/formatter on modified files `main.py` and `tests/test_main_cli_args.py` to check styling.
