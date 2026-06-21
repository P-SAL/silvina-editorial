from unittest import TestCase
from unittest.mock import patch

from src.domain.exceptions.language_model_errors import LanguageModelUnavailable
from src.infrastructure.adapters.llm_generator.ollama_generator_adapter import (
    OllamaGeneratorAdapter,
)


class TestOllamaGeneratorAdapter(TestCase):
    def setUp(self):
        self.adapter = OllamaGeneratorAdapter(
            model_name="llama3-gradient:8b-instruct-1048k-q4_K_M",
            base_url="http://localhost:11434",
        )

    @patch("src.infrastructure.adapters.llm_generator.ollama_generator_adapter.ollama.generate")
    def test_generate_returns_stripped_response_text(self, mock_generate):
        mock_generate.return_value = {"response": "  some text  "}

        result = self.adapter.generate("prompt")

        self.assertEqual(result, "some text")

    @patch("src.infrastructure.adapters.llm_generator.ollama_generator_adapter.ollama.generate")
    def test_generate_raises_language_model_unavailable_on_backend_failure(self, mock_generate):
        mock_generate.side_effect = ConnectionError("backend unreachable")

        with self.assertRaises(LanguageModelUnavailable):
            self.adapter.generate("prompt")
