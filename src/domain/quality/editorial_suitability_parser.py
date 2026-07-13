from re import IGNORECASE, Pattern, compile as re_compile

_VEREDICTO_PATTERN = re_compile(r"VEREDICTO:\s*(.+)", IGNORECASE)
_CONTRIBUCION_PATTERN = re_compile(r"CONTRIBUCI[OÓ]N:\s*(.+)", IGNORECASE)
_LINEAS_PATTERN = re_compile(r"L[IÍ]NEAS:\s*(.+)", IGNORECASE)
_JUSTIFICACION_PATTERN = re_compile(r"JUSTIFICACI[OÓ]N:\s*(.+)", IGNORECASE)
_SENTENCE_END_PATTERN = re_compile(r"[.!?]")

_CONTRIBUTION_VERDICTS = ("NO SUSTENTADA", "PARCIAL", "SUSTENTADA")
_ALIGNMENT_VERDICTS = ("NO ALINEADO", "PARCIALMENTE ALINEADO", "ALINEADO")

_PHRASE_MAX_LENGTH = 120
_OBSERVATION_MAX_LENGTH = 120
_JUSTIFICATION_MAX_LENGTH = 120
_LINES_MAX_LENGTH = 80

_NO_SUSTENTADA_OBSERVATION = "Sin contribución observada o declarada."
_PARCIAL_OBSERVATION = "Contribución declarada pero no suficientemente sustentada."
_SUSTENTADA_OBSERVATION_FALLBACK = "Contribución sustentada."


class EditorialSuitabilityParser:
    """Stateless parser that extracts verdicts and justifications from raw LLM text."""

    def parse_contribution(self, text: str) -> tuple[str, str, str]:
        """Return (contribution_verdict, contribution_phrase, contribution_observation)."""
        verdict = self._extract_verdict(text, _CONTRIBUTION_VERDICTS)
        phrase = self._truncate_field(
            self._extract_field(text, _CONTRIBUCION_PATTERN), _PHRASE_MAX_LENGTH
        )
        observation = self._build_contribution_observation(verdict, phrase)
        return verdict, phrase, observation

    def parse_alignment(self, text: str) -> tuple[str, str, str]:
        """Return (alignment_verdict, alignment_lines, alignment_justification)."""
        verdict = self._extract_verdict(text, _ALIGNMENT_VERDICTS)
        lines = self._truncate_field(self._extract_field(text, _LINEAS_PATTERN), _LINES_MAX_LENGTH)
        justification = self._truncate_field(
            self._extract_field(text, _JUSTIFICACION_PATTERN), _JUSTIFICATION_MAX_LENGTH
        )
        return verdict, lines, justification

    def _build_contribution_observation(self, verdict: str, phrase: str) -> str:
        if verdict == "NO SUSTENTADA":
            return _NO_SUSTENTADA_OBSERVATION
        if verdict == "PARCIAL":
            return _PARCIAL_OBSERVATION
        if not phrase:
            return _SUSTENTADA_OBSERVATION_FALLBACK
        return self._truncate_field(f"Contribución sustentada — {phrase}", _OBSERVATION_MAX_LENGTH)

    def _extract_verdict(self, text: str, candidates: tuple[str, ...]) -> str:
        match = _VEREDICTO_PATTERN.search(text)
        raw = match.group(1).upper() if match else ""
        for candidate in candidates:
            if candidate in raw:
                return candidate
        return candidates[0]

    def _extract_field(self, text: str, pattern: Pattern) -> str:
        match = pattern.search(text)
        return match.group(1).strip() if match else ""

    def _truncate_field(self, raw: str, max_length: int) -> str:
        if not raw:
            return raw
        sentence = self._extract_first_sentence(raw)
        if len(sentence) < max_length:
            return sentence
        return self._truncate_to_word_boundary(sentence, max_length)

    def _extract_first_sentence(self, text: str) -> str:
        match = _SENTENCE_END_PATTERN.search(text)
        if match:
            return text[: match.start() + 1].strip()
        return text.strip()

    def _truncate_to_word_boundary(self, text: str, max_length: int) -> str:
        limit = max_length - 2
        truncated = text[:limit]
        last_space = truncated.rfind(" ")
        if last_space > 0:
            truncated = truncated[:last_space]
        return f"{truncated.rstrip(' .,;:')}…"
