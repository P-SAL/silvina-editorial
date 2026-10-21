from src.domain.ports.llm_generator_port import LlmGeneratorPort


class FakeLlmGeneratorAdapter(LlmGeneratorPort):
    """Test double for LlmGeneratorPort that returns canned responses in order."""

    def __init__(self, responses: list[str]) -> None:
        self._responses = responses
        self.call_count = 0
        self.received_prompts: list[str] = []
        self.received_options: list[dict | None] = []

    def generate(self, prompt: str, options: dict | None = None) -> str:
        """Return the next canned response and record the received prompt and options."""
        self.received_prompts.append(prompt)
        self.received_options.append(options)
        response = self._responses[self.call_count]
        self.call_count += 1
        return response
