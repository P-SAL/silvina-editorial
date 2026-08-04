from src.domain.document.character_count_port import CharacterCountPort
from src.domain.dtos.character_count_dto import CharacterCountDTO


class FakeCharacterCountPort(CharacterCountPort):
    """Test double for CharacterCountPort with configurable CharacterCountDTO, None return, or exception."""

    def __init__(
        self,
        result: CharacterCountDTO | None = None,
        error: Exception | None = None,
    ) -> None:
        self._result = result
        self._error = error

    def count(self, docx_path: str) -> CharacterCountDTO | None:
        """Return the configured result, raise the configured exception, or return None."""
        if self._error is not None:
            raise self._error
        return self._result
