from src.domain.exceptions.base_src_error import SrcBaseWarning


class QualityAnalysisFailed(SrcBaseWarning):
    """Raised when quality analysis cannot be completed."""

    MESSAGE = "The quality analysis could not be completed."
