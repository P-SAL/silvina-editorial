import re
from datetime import datetime

from src.domain.dtos.document_content_dto import DocumentContentDTO


class ReferenceSignalDetector:
    """Domain service that detects reference-count and reference-recency signals (S2a, S2b)."""

    _YEAR_PATTERN = re.compile(r"\b((?:19|20)\d{2})\b")

    def __init__(
        self,
        minimum_reference_count: int = 12,
        recent_reference_year_offset: int = 4,
        minimum_recent_reference_ratio: float = 0.5,
    ) -> None:
        self._minimum_reference_count = minimum_reference_count
        self._recent_reference_year_offset = recent_reference_year_offset
        self._minimum_recent_reference_ratio = minimum_recent_reference_ratio

    def has_recent_majority(self, document_content: DocumentContentDTO) -> bool:
        """Return whether at least half of the references are recent (signal S2b)."""
        references = document_content.references
        if not references:
            return False

        recent_threshold = datetime.now().year - self._recent_reference_year_offset
        recent_count = 0
        for reference in references:
            years = [int(year) for year in self._YEAR_PATTERN.findall(reference.text)]
            if years and max(years) >= recent_threshold:
                recent_count += 1

        return (recent_count / len(references)) >= self._minimum_recent_reference_ratio

    def has_sufficient_count(self, document_content: DocumentContentDTO) -> bool:
        """Return whether the document has enough references (signal S2a)."""
        return len(document_content.references) >= self._minimum_reference_count
