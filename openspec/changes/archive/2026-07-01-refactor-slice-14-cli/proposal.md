# Proposal: Refactor Slice 14 CLI (refactor-slice-14-cli)

## 1. Intent
Refactor `main.py` to use Hexagonal Architecture Use Cases (`AnalyzeDocumentUseCase` and `ExportReportUseCase`) and their wirings, while preserving exact console output styling and compatibility with existing E2E tests and `gradio_app.py`.

## 2. Scope
- Refactor `main.py` to be a lightweight driver.
- Implement DTO-to-legacy dictionary mapping for compatibility.
- Configure report paths to be configurable (as arguments/options to main) with current paths (same directory as input) as defaults.
- Fail immediately if Ollama/LLM backend connection fails.
- Let validation/file checks flow from the domain and catch them.
- Maintain `main_legacy.py` (already created) as a copy of the old `main.py`.

## 3. Approach
- Instantiate `AnalyzeDocumentUseCase` using `AnalyzeDocumentUseCaseWiring`.
- Instantiate `ExportReportUseCase` using `ExportReportWiring`.
- Refactor `SilvinaEditorialAssistant` to delegate to Use Cases and map output via DTO-to-legacy mapping helper.
- Handle exceptions inheriting from `BaseSrcError` cleanly, exit with status codes (0 for success/user-interrupt, 1 for domain/generic errors, 2 for arguments/value errors).

## 4. Capabilities
- **New Capabilities**: None
- **Modified Capabilities**: None (pure CLI wiring refactoring)

## 5. Affected Areas
- [main.py](file:///E:/Python/silvina-editorial/main.py)

## 6. Rollback Plan
- Copy [main_legacy.py](file:///E:/Python/silvina-editorial/main_legacy.py) back to [main.py](file:///E:/Python/silvina-editorial/main.py).
