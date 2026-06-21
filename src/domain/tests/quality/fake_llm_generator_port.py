class FakeLlmGeneratorPort:
    def __init__(self, responses: list[str]) -> None:
        self._responses = responses
        self.call_count = 0
        self.received_prompts: list[str] = []

    def generate(self, prompt: str) -> str:
        self.received_prompts.append(prompt)
        response = self._responses[self.call_count]
        self.call_count += 1
        return response
