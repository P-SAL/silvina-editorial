from enum import Enum


class ValidationStatus(Enum):
    """Status of validation checks."""

    PASSED = "passed"
    FAILED = "failed"
    WARNING = "warning"
    NOT_APPLICABLE = "not_applicable"
