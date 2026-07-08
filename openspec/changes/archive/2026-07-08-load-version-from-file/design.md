# Design: Load version from file

## Technical Approach

Decouple application versioning from environment variables by removing `SILVINA_VERSION` from env configuration files (`.env`/`.env.example`) and resolving it dynamically from a root `version.txt` file. To prevent test suite breakage, fallback to environment variable resolution when `TESTING` is set.

## Architecture Decisions

| Option | Tradeoff | Decision |
|--------|----------|----------|
| Dynamic file reading | Adds file I/O overhead on config load; decouples version from env config. | Read `version.txt` dynamically at instantiation. |
| Testing fallback | Bypasses file system dependency in tests; introduces conditional path. | Use `TESTING` environment variable fallback to support lightweight testing. |

## Data Flow

```
[version.txt] ──────────┐
(Production)            │
                        ▼
                  [EnvConfig] ──(silvina_version)──→ [DocxReportAdapter]
                        ▲
(Testing)               │
[SILVINA_VERSION] ──────┘ (when TESTING=True)
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `version.txt` | Create | Root version file with plain string `0.95`. |
| `.env` | Modify | Remove `SILVINA_VERSION`. |
| `.env.example` | Modify | Remove `SILVINA_VERSION`. |
| `src/infrastructure/env_config.py` | Modify | Resolve version dynamically from `version.txt` or fallback in testing mode. |
| `conftest.py` | Modify | Set `os.environ["TESTING"] = "True"` at startup. |
| `src/infrastructure/tests/test_env_config.py` | Modify | Mock file resolution and verify new version loading scenarios. |
| `src/infrastructure/tests/test_export_report_wiring.py` | Modify | Include `TESTING: "True"` in environment patch for wiring tests. |
| `openspec/specs/analyze-document/spec.md` | Modify | Update environment variable configuration table. |
| `openspec/TECHNICAL_DEBT.md` | Modify | Move technical debt item 7 to resolved list. |

## Interfaces / Contracts

No new public interfaces or contracts. `EnvConfig.silvina_version` remains a public `str` property.

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | Production file loading | Mock file existence and content to assert successful version assignment. |
| Unit | Production missing file | Mock file absence to assert raising of `FileNotFoundError`. |
| Unit | Testing fallback | Patch environment with `TESTING="True"` and assert resolution from env. |
| Integration | Wiring version resolution | Patch `TESTING="True"` in wiring tests to verify version is correctly injected into adapters. |

## Migration / Rollout

No migration required. Ensure `version.txt` is committed to version control and distributed in deployments.
