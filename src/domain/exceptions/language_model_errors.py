from src.domain.exceptions.base_src_error import SrcBaseWarning


class LanguageModelUnavailable(SrcBaseWarning):
    """Raised when the language model backend is unavailable."""

    MESSAGE = "The language model backend is unavailable."
