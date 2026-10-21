from src.domain.exceptions.base_src_error import BaseSrcError, SrcBaseWarning


class CountError(BaseSrcError):
    """Base class for all character count exceptions."""


class CharacterCountUnavailable(SrcBaseWarning):
    """Raised when COM-based character counting fails."""

    MESSAGE = "Character count via Word COM is unavailable."
