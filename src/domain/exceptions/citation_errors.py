from src.domain.exceptions.base_src_error import SrcBaseWarning


class CitationParsingFailed(SrcBaseWarning):
    """Raised when a citation cannot be parsed from the document."""

    MESSAGE = "The citation could not be parsed."
