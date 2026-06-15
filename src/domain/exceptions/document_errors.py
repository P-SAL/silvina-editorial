from src.domain.exceptions.base_src_error import SrcBaseNotFound, SrcBaseWarning


class DocumentNotFound(SrcBaseNotFound):
    """Raised when a document file cannot be located."""

    MESSAGE = "The document file could not be found."


class DocumentEmpty(SrcBaseWarning):
    """Raised when a document has no readable content."""

    MESSAGE = "The document has no readable content."


class DocumentUnreadable(SrcBaseWarning):
    """Raised when a document cannot be read or parsed."""

    MESSAGE = "The document could not be read."
