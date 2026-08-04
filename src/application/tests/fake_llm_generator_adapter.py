from src.domain.ports.llm_generator_port import LlmGeneratorPort


class FakeLlmGeneratorAdapterForTest(LlmGeneratorPort):
    def generate(self, prompt: str, options: dict | None = None) -> str:
        return (
            "**Claridad** [Puntuación: 8/10] Texto claro y bien estructurado en general.\n"
            "**Coherencia** [Puntuación: 8/10] Las ideas se conectan de forma coherente entre sí.\n"
            "**Argumentación** [Puntuación: 8/10] Los argumentos están bien fundamentados y son sólidos.\n"
            "**Conclusiones** [Puntuación: 8/10] Las conclusiones se derivan correctamente del análisis.\n"
        )
