# Design: Resolve Bare Imports

## Technical Approach

We will refactor standard library imports in `main.py` and `tests/test_main_cli_args.py` to adhere to the project's Clean Hexagonal Architecture conventions. Instead of using bare module imports (e.g., `import os`), specific names will be imported and used directly. No functionality or behavior changes will be introduced.

This maps directly to the proposal's approach, targeting only `main.py` and `tests/test_main_cli_args.py`.

## Architecture Decisions

### Decision: Refactoring Bare Imports Style

| Option | Tradeoff | Decision |
|--------|----------|----------|
| **Direct Import Refactoring** | Very low risk, direct mapping to specific names, strictly follows clean architecture guidelines. | **Chosen**. Standardizes import style without changing code execution path. |
| **Refactoring path logic to pathlib** | Slightly cleaner code syntax using pathlib, but increases code diff size and regression risk. | **Rejected**. The goal is focused on resolving bare imports. |

### Decision: Mock Patching Method in Tests

| Option | Tradeoff | Decision |
|--------|----------|----------|
| **`patch("sys.argv", ...)`** | Standard string-based patch, removes the need to import `sys` module in tests completely. | **Chosen**. Simplifies imports and uses standard mock patching. |
| **`patch.object(sys, "argv", ...)`** | Requires importing `sys` solely to access the `argv` attribute. | **Rejected**. Unnecessarily keeps a dependency on the `sys` module in the test file. |

## Data Flow

No changes to the data flow or the application workflow.

```
[CLI Invocation] ──→ [main.py (main)] ──→ [SilvinaEditorialAssistant]
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| [main.py](file:///E:/Python/silvina-editorial/main.py) | Modify | Import specific names from standard libraries (`argparse`, `re`, `sys`, `os.path`, `traceback`, `json`). Update references in the code. |
| [tests/test_main_cli_args.py](file:///E:/Python/silvina-editorial/tests/test_main_cli_args.py) | Modify | Import specific names from standard libraries (`io`, `os.path`, `sys`, `unittest`). Update references and replace `patch.object(sys, "argv", ...)` with `patch("sys.argv", ...)`. |

## Interfaces / Contracts

No new interfaces or API contracts are introduced. The standard library imports are updated as follows:

In [main.py](file:///E:/Python/silvina-editorial/main.py):
```python
from argparse import ArgumentParser
from json import dump
from os.path import exists, join
from pathlib import Path
from re import sub
from sys import exit, path, stderr, stdout
from traceback import print_exc
```

In [tests/test_main_cli_args.py](file:///E:/Python/silvina-editorial/tests/test_main_cli_args.py):
```python
from io import StringIO
from os.path import dirname, join
from sys import path
from unittest import TestCase, main
from unittest.mock import patch
```

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | CLI arguments and main entrypoint exit codes. | Run the existing test suite: `.venv/Scripts/python.exe -m unittest tests/test_main_cli_args.py` |

## Migration / Rollout

No migration required.

## Open Questions

None.
