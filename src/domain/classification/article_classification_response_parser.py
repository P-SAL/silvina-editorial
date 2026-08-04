import re


class ArticleClassificationResponseParser:
    """Parses the combined S4/S5/S6 LLM response into 3 boolean signals."""

    _RESEARCH_INTENT_PATTERN = re.compile(r"S4\s*:\s*SI")
    _EVIDENCE_BASED_CONTRIBUTION_PATTERN = re.compile(r"S5\s*:\s*SI")
    _THEORETICAL_JUSTIFICATION_PATTERN = re.compile(r"S6\s*:\s*SI")

    def parse(self, response_text: str) -> tuple[bool, bool, bool]:
        """Return (s4, s5, s6) extracted from the LLM's free-text yes/no response."""
        response_upper = response_text.strip().upper()
        s4 = bool(self._RESEARCH_INTENT_PATTERN.search(response_upper))
        s5 = bool(self._EVIDENCE_BASED_CONTRIBUTION_PATTERN.search(response_upper))
        s6 = bool(self._THEORETICAL_JUSTIFICATION_PATTERN.search(response_upper))
        return s4, s5, s6
