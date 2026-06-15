from src.domain.exceptions.base_src_error import BaseSrcError


class LanguageModelError(BaseSrcError):
    """Base class for all language model-related exceptions."""


class LanguageModelUnavailable(LanguageModelError):
    """Raised when the language model backend is unavailable."""

    MESSAGE = "The language model backend is unavailable."
