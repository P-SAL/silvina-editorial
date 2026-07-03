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

    @patch("src.infrastructure.adapters.llm_generator.ollama_generator_adapter.ollama.Client")
    def test_generate_returns_stripped_response_text(self, mock_client_class):
        mock_client = mock_client_class.return_value
        mock_client.generate.return_value = {"response": "  some text  "}

        result = self.adapter.generate(prompt="prompt")

        self.assertEqual(result, "some text")

    @patch("src.infrastructure.adapters.llm_generator.ollama_generator_adapter.ollama.Client")
    def test_generate_raises_language_model_unavailable_on_backend_failure(self, mock_client_class):
        mock_client = mock_client_class.return_value
        mock_client.generate.side_effect = ConnectionError("backend unreachable")

        with self.assertRaises(LanguageModelUnavailable):
            self.adapter.generate(prompt="prompt")

    @patch("src.infrastructure.adapters.llm_generator.ollama_generator_adapter.ollama.Client")
    def test_generate_instantiates_client_with_configured_base_url(self, mock_client_class):
        mock_client = mock_client_class.return_value
        mock_client.generate.return_value = {"response": "some text"}

        self.adapter.generate(prompt="prompt")

        mock_client_class.assert_called_once_with(host="http://localhost:11434")

    @patch("src.infrastructure.adapters.llm_generator.ollama_generator_adapter.ollama.Client")
    def test_generate_forwards_options_dict_to_ollama_generate(self, mock_client_class):
        mock_client = mock_client_class.return_value
        mock_client.generate.return_value = {"response": "some text"}

        self.adapter.generate(prompt="prompt", options={"temperature": 0.1, "num_predict": 300})

        mock_client.generate.assert_called_once_with(
            model="llama3-gradient:8b-instruct-1048k-q4_K_M",
            prompt="prompt",
            options={"temperature": 0.1, "num_predict": 300},
        )

    @patch("src.infrastructure.adapters.llm_generator.ollama_generator_adapter.ollama.Client")
    def test_generate_without_options_argument_preserves_prior_behavior(self, mock_client_class):
        mock_client = mock_client_class.return_value
        mock_client.generate.return_value = {"response": "some text"}

        self.adapter.generate(prompt="prompt")

        mock_client.generate.assert_called_once_with(
            model="llama3-gradient:8b-instruct-1048k-q4_K_M",
            prompt="prompt",
            options=None,
        )
