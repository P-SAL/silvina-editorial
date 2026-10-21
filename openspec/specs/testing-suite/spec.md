# Testing Suite Specification

## Purpose

This specification defines the requirements and scenarios for the automated testing suite of the Silvina Editorial Assistant. It covers unit testing of validators, integration testing of file parsers using docx fixtures, mocking of external services (COM, Ollama, LanguageTool), and end-to-end testing of both CLI and UI.

## Requirements

### Requirement: Deterministic Validator Unit Testing

The test suite SHALL verify the core business logic in `StructureValidator` and `APAValidator` deterministically, running entirely in-memory without making external network or system calls.

#### Scenario: Valid Document Structure Verification
- GIVEN a document representation with all required sections present
- WHEN `StructureValidator` analyzes the document structure
- THEN the validator returns a successful check status with no structure errors

#### Scenario: Invalid Citation Formatting Detection
- GIVEN a citation that does not comply with APA style guidelines
- WHEN `APAValidator` inspects the citation format
- THEN the validator flags the citation as incorrect and returns specific style warnings

### Requirement: File Parser Integration Verification

The integration test suite MUST verify the parsing accuracy of `WordReader`, `CitationParser`, and `ReferenceParser` using actual `.docx` files as test fixtures, including the document `capacidades_razonamiento_emergente_LLMs.docx`.

#### Scenario: Document and Reference Parsing
- GIVEN a valid `.docx` test fixture containing structured text and citations
- WHEN the parser components extract content, citations, and references
- THEN the parsed references match the expected source structures exactly

#### Scenario: Empty or Corrupted File Parsing
- GIVEN a corrupted or empty file fixture
- WHEN `WordReader` attempts to read the file
- THEN the reader raises an appropriate validation exception without crashing

### Requirement: External Dependency Mocking and Stubbing

The testing framework MUST mock or stub all external system and API dependencies—specifically `win32com` COM interfaces, Ollama clients, and LanguageTool Java services—to allow successful test execution on platforms lacking these dependencies, such as CI runners.

#### Scenario: CI Compatibility with Missing External Services
- GIVEN a test environment without MS Word, local Ollama endpoints, or Java runtimes
- WHEN the full test suite runs
- THEN all tests pass successfully without trying to initiate live external services

### Requirement: CLI Workflow End-to-End Validation

The testing suite MUST execute the `main.py` CLI workflow end-to-end, utilizing mocked API adapters, and confirm that correct report artifacts are written to the filesystem.

#### Scenario: Successful CLI Execution and Report Generation
- GIVEN a valid test configuration and mock data source
- WHEN the user executes `main.py` through the CLI wrapper with arguments
- THEN the system exits with a zero exit status and generates a markdown report in the designated output path

### Requirement: UI Interaction End-to-End Validation

The test suite MUST support executing `gradio_app.py` in test mode using the native Gradio test client to simulate user uploads and verify interface responses.

#### Scenario: File Upload UI Simulation
- GIVEN the Gradio test client connected to the application in test mode
- WHEN a test user uploads a `.docx` file through the interface
- THEN the test client receives a successful response and reports validation results on the UI state

#### Scenario: Upload of Unsupported Format
- GIVEN the Gradio test client connected to the application in test mode
- WHEN a test user uploads a file with an unsupported file extension
- THEN the interface displays a user-friendly error message indicating the invalid format

## Delta: fix-source-defects

### Requirement: CitationMatcher Report Generation with Orphaned Citations (REQ-EXT-1)

The `generate_report()` method in `CitationMatcher` SHALL handle orphaned citations gracefully, including severity labeling in the report output without crashing.

#### Scenario: Report generation with orphaned citations
- GIVEN a `CitationMatcher` instance with citations that have no matching references
- WHEN `generate_report()` is called
- THEN the method returns a report string that includes a severity label (e.g., 'critical') and does not raise an exception

### Requirement: APAValidator ET_AL Format Detection and Comment Cleanup (REQ-EXT-2)

The `APAValidator` SHALL correctly detect "et. al." (with inner period) as `ET_AL_FORMAT_ERROR`. The test suite SHALL contain no stale "known limitation" comments that document incomplete validator behavior.

#### Scenario: Detect et. al. with inner period
- GIVEN a citation string in the format "(Author et. al., Year)"
- WHEN `APAValidator.validate_citation()` is called
- THEN the validator returns an error list containing `ErrorType.ET_AL_FORMAT_ERROR`

#### Scenario: No stale limitation comments
- GIVEN the test suite source files
- WHEN tests are examined
- THEN no comments documenting known limitations in ET_AL format detection exist

### Requirement: Developer Setup Script for Virtual Environment (REQ-W4)

A `setup.bat` script SHALL be present at the repository root to enable developers on Windows to bootstrap the virtual environment and install dependencies in a single command.

#### Scenario: setup.bat creates and provisions virtual environment
- GIVEN a clean clone of the repository on Windows
- WHEN a developer runs `setup.bat` from the repository root
- THEN the script creates a `.venv` directory, installs dependencies from `requirements.txt`, and exits with status code 0

#### Scenario: setup.bat error handling
- GIVEN the `setup.bat` script encounters an error during venv creation or pip install
- WHEN the error occurs
- THEN the script displays an appropriate error message and exits with a non-zero status code
