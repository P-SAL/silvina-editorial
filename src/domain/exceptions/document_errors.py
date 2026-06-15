from src.domain.exceptions.base_src_error import BaseSrcError


class DocumentError(BaseSrcError):
    """Base class for all document-related exceptions."""


class DocumentNotFound(DocumentError):
    """Raised when a document file cannot be located."""

    MESSAGE = "The document file could not be found."


class DocumentEmpty(DocumentError):
    """Raised when a document has no readable content."""

    MESSAGE = "The document has no readable content."


class DocumentUnreadable(DocumentError):
    """Raised when a document cannot be read or parsed."""

    MESSAGE = "The document could not be read."
