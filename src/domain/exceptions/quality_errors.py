from src.domain.exceptions.base_src_error import BaseSrcError


class QualityError(BaseSrcError):
    """Base class for all quality-related exceptions."""


class QualityAnalysisFailed(QualityError):
    """Raised when quality analysis cannot be completed."""

    MESSAGE = "The quality analysis could not be completed."
