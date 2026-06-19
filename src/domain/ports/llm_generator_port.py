from typing import Protocol


class LlmGeneratorPort(Protocol):
    """Capability to generate text from a prompt via a language model backend."""

    def generate(self, prompt: str) -> str:
        """Return the generated text for the given prompt."""
        ...
