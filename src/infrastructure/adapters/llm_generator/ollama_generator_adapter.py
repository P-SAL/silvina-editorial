import ollama

from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler
from src.domain.exceptions.language_model_errors import LanguageModelUnavailable
from src.domain.ports.llm_generator_port import LlmGeneratorPort


class OllamaGeneratorAdapter(LlmGeneratorPort):
    """Generates text via a local Ollama backend."""

    def __init__(self, model_name: str, base_url: str) -> None:
        self._model_name = model_name
        self._base_url = base_url

    @generic_error_handler
    def generate(self, prompt: str, options: dict | None = None) -> str:
        """Return Ollama's generated text for the given prompt."""
        try:
            client = ollama.Client(host=self._base_url)
            response = client.generate(model=self._model_name, prompt=prompt, options=options)
        except (ollama.RequestError, ollama.ResponseError, ConnectionError) as exc:
            raise LanguageModelUnavailable() from exc
        return response.get("response", "").strip()
