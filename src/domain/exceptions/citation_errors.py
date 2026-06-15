from src.domain.exceptions.base_src_error import BaseSrcError


class CitationError(BaseSrcError):
    """Base class for all citation-related exceptions."""


class CitationParsingFailed(CitationError):
    """Raised when a citation cannot be parsed from the document."""

    MESSAGE = "The citation could not be parsed."
