from src.domain.exceptions.base_src_error import BaseSrcError


class ClassificationError(BaseSrcError):
    """Base class for all classification-related exceptions."""


class ClassificationFailed(ClassificationError):
    """Raised when article classification cannot be completed."""

    MESSAGE = "The article classification could not be completed."
