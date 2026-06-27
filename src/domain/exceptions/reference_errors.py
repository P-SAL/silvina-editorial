from src.domain.exceptions.base_src_error import BaseSrcError


class ReferenceExtractionError(BaseSrcError):
    """Base class for all reference-related exceptions."""


class ReferenceParsingFailed(ReferenceExtractionError):
    """Raised when a reference cannot be parsed from the document."""

    MESSAGE = "The reference could not be parsed."
