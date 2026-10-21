from src.domain.exceptions.base_src_error import SrcBaseWarning


class GrammarError(SrcBaseWarning):
    """Base class for all grammar check exceptions."""


class GrammarCheckUnavailable(GrammarError):
    """Raised when the grammar checker backend is unavailable."""

    MESSAGE = "The grammar check service is unavailable."
