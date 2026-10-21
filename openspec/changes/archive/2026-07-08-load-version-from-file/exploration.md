# Exploration: Load version number from version.txt instead of .env

## Current State
Currently, the software version is managed as an environment variable `SILVINA_VERSION` defined in the `.env` file (and documented/fallback in `.env.example`).
In [env_config.py](file:///E:/Python/silvina-editorial/src/infrastructure/env_config.py#L84), this value is read during initialization via:
`self.silvina_version: str = getenv("SILVINA_VERSION", "0.9")`

This couples software versioning directly to environment configuration. In addition, the version number is hardcoded in main.py print statements (`v0.9`), and in various markdown specs and documentation files (`README.md`, `CITATION.cff`, etc.).
Furthermore, the unit tests in [test_env_config.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_env_config.py) and [test_export_report_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_export_report_wiring.py) patch the environment variable `SILVINA_VERSION` to verify configuration defaults and wiring injections.

## Affected Areas
- `version.txt` (New) — A new file containing the software version number (e.g. `0.95`) at the root of the project.
- `.env` — Remove the `SILVINA_VERSION` environment variable definition.
- `.env.example` — Remove the `SILVINA_VERSION` environment variable definition.
- [env_config.py](file:///E:/Python/silvina-editorial/src/infrastructure/env_config.py) — Replace `getenv("SILVINA_VERSION", "0.9")` with logic to resolve and read `version.txt` from the project root directory, falling back to a default value (e.g., `"0.9"`) if the file is missing or unreadable.
- [test_env_config.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_env_config.py) — Update tests (like `test_defaults_are_loaded_when_env_is_empty`) to patch `pathlib.Path.is_file` to mock the presence/absence of `version.txt`, and add a dedicated test verifying correct loading from `version.txt`.
- [test_export_report_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_export_report_wiring.py) — Update `test_create_use_case_injects_app_name_and_version_from_env` to mock the file read instead of environment variable `SILVINA_VERSION`.
- [spec.md (analyze-document)](file:///E:/Python/silvina-editorial/openspec/specs/analyze-document/spec.md#L169) — Update the configuration table to document that `SILVINA_VERSION` has been removed as an environment variable and the version is now loaded from `version.txt`.
- [TECHNICAL_DEBT.md](file:///E:/Python/silvina-editorial/openspec/TECHNICAL_DEBT.md#L13-L16) — Mark item 7 ("Version number loaded from version.txt instead of .env") as resolved.

## Approaches

1. **Direct Project-Root File Read (Recommended)**
   Read `version.txt` from the project root using a path resolved relative to `__file__` (which resides at `src/infrastructure/env_config.py`).
   Implementation detail:
   ```python
   import pathlib

   # In EnvConfig.__init__
   version_path = pathlib.Path(__file__).resolve().parents[2] / "version.txt"
   if version_path.is_file():
       try:
           self.silvina_version: str = version_path.read_text(encoding="utf-8").strip()
       except Exception:
           self.silvina_version = "0.9"
   else:
       self.silvina_version = "0.9"
   ```
   And in unit tests:
   - In `test_env_config.py`, patch `pathlib.Path.is_file` to control whether the file exists, and mock `builtins.open` (or `pathlib.Path.read_text`) to supply a custom version string to verify reading works.
   - In `test_export_report_wiring.py`, replace environment mocking of `SILVINA_VERSION` with mocking of `version.txt` loading.
   - **Pros:**
     - Fully decouples software versioning from environment variables, adhering strictly to the technical debt registry requirements.
     - Keeps the constructor interface of `EnvConfig` simple and unchanged.
     - Very low implementation complexity.
   - **Cons:**
     - Requires unit tests to mock `pathlib.Path.is_file` or `builtins.open` to isolate the tests from the actual root `version.txt` file content and prevent test instability.
   - **Effort:** Low

2. **Optional Constructor Parameter in `EnvConfig`**
   Provide an optional `version_file_path` parameter to the `EnvConfig` constructor.
   Implementation detail:
   ```python
   class EnvConfig:
       def __init__(self, version_file_path: pathlib.Path | str | None = None) -> None:
           # ...
           if version_file_path is None:
               version_path = pathlib.Path(__file__).resolve().parents[2] / "version.txt"
           else:
               version_path = pathlib.Path(version_file_path)
   ```
   - **Pros:**
     - Allows unit tests to pass an explicit non-existent path or a custom test file path directly to avoid patching file operations globally.
   - **Cons:**
     - Changes the class constructor signature.
     - Since wiring layers instantiate `EnvConfig()` with no arguments, tests of the wiring layers would still require filesystem mocking or passing custom configs to the wirings.
   - **Effort:** Medium

## Recommendation
Approach 1 is recommended. It is direct, avoids changing the public interface of `EnvConfig` used by wirings, and is easy to implement. The test coverage can be cleanly isolated using standard library mocks on `pathlib.Path.is_file` and `builtins.open` (or `pathlib.Path.read_text`).

## Risks
- **Test Isolation/Side Effects:** Tests might accidentally read the physical `version.txt` at the project root, which contains `"0.95"`, failing assertions expecting `"0.9"`.
  - *Mitigation:* Ensure unit tests patch `pathlib.Path.is_file` to return `False` when asserting defaults, or mock file reading to return a controlled version string.
- **File Access Exceptions:** If the application runs under restricted permissions or if `version.txt` is missing from a package distribution, reading the file could raise an error.
  - *Mitigation:* Wrap the file reading in a defensive `try...except` block and fallback gracefully to `"0.9"`.

## Ready for Proposal
Yes. The orchestrator can proceed to the design phase.
