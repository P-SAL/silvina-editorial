# Proposal: Resolve Explicit Keyword Arguments

## Intent
Enforce consistency and readability by migrating positional argument calls to explicit keyword argument calls in custom code. This resolves the technical debt outlined in `TECHNICAL_DEBT.md` Item 1, aligning legacy modules with the project convention.

## Scope

### In Scope
- Refactor all custom function and method calls in the application and infrastructure layers to use explicit keyword arguments.
- Cover all custom code in `src/application/`, `src/infrastructure/wirings/`, and `src/infrastructure/adapters/`.
- Update corresponding unit tests in `src/application/tests/` and `src/infrastructure/tests/` to use keyword arguments for custom class/method calls.

### Out of Scope
- Standard library calls (e.g., `os.path.join`, `open`).
- Third-party library calls (e.g., `python-docx`, `gradio`).
- Standard unit test assertion calls (e.g., `self.assertEqual`, `self.assertRaises`).

## Capabilities

### New Capabilities
None

### Modified Capabilities
None

## Approach
Implement Option 2: Full Audit and Alignment of All Internal and External Calls.
1. Audit and update all target files to ensure custom calls pass arguments by name (e.g. `use_case.execute(document_path=...)`).
2. Run standard formatters and linters after updating code to ensure compliance.
3. Keep parameter signature names unchanged to avoid introducing breaking changes.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/application/` | Modified | Update execution calls to sub-use cases and interfaces. |
| `src/infrastructure/wirings/` | Modified | Update instantiation and helper method calls. |
| `src/infrastructure/adapters/` | Modified | Update internal helper calls and interface implementations. |
| `src/application/tests/` | Modified | Update test execution calls. |
| `src/infrastructure/tests/` | Modified | Update test execution calls. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Runtime argument name mismatch | Low | Rely on exact matching of existing parameter names, and run tests. |
| Regression in test files | Low | Run complete test suite (`pytest` and `python -m unittest`) before and after changes. |

## Rollback Plan
Revert changes using git: `git checkout -- src/` or revert the specific commit introducing the change.

## Dependencies
None

## Success Criteria
- [ ] All custom function and method calls in the audited directories use explicit keyword arguments.
- [ ] No regression is introduced; 100% of existing tests pass successfully.
