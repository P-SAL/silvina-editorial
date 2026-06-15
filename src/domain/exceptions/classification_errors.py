from src.domain.exceptions.base_src_error import SrcBaseWarning


class ClassificationFailed(SrcBaseWarning):
    """Raised when article classification cannot be completed."""

    MESSAGE = "The article classification could not be completed."
