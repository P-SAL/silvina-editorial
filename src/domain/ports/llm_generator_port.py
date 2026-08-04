from abc import ABC, abstractmethod


class LlmGeneratorPort(ABC):
    """Capability to generate text from a prompt via a language model backend."""

    @abstractmethod
    def generate(self, prompt: str, options: dict | None = None) -> str:
        """Return the generated text for the given prompt."""
