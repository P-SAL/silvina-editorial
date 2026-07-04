from dotenv import load_dotenv
from os import getenv

from src.application.analyze_quality_use_case import AnalyzeQualityUseCase
from src.domain.dtos.quality_level_thresholds_dto import QualityLevelThresholdsDTO
from src.domain.ports.llm_generator_port import LlmGeneratorPort
from src.domain.quality.quality_analyzer import QualityAnalyzer
from src.domain.quality.quality_level_resolver import QualityLevelResolver
from src.domain.quality.quality_response_parser import QualityResponseParser
from src.domain.quality.quality_text_sampler import QualityTextSampler
from src.infrastructure.adapters.llm_generator.ollama_generator_adapter import (
    OllamaGeneratorAdapter,
)
from src.infrastructure.resources.prompts.quality import PROMPTS_DIR
from src.infrastructure.resources.text_resource_loader import read_text_resource

load_dotenv()


class AnalyzeQualityUseCaseWiring:
    """Factory for building a ready-to-use AnalyzeQualityUseCase."""

    def create_use_case(self) -> AnalyzeQualityUseCase:
        return AnalyzeQualityUseCase(analyzer=self._get_quality_analyzer())

    def _get_llm_generator(self) -> LlmGeneratorPort:
        model_name = getenv("OLLAMA_MODEL_NAME", "llama3-gradient:8b-instruct-1048k-q4_K_M")
        base_url = getenv("OLLAMA_BASE_URL", "http://localhost:11434")
        return OllamaGeneratorAdapter(model_name=model_name, base_url=base_url)

    def _get_quality_analyzer(self) -> QualityAnalyzer:
        return QualityAnalyzer(
            llm_generator=self._get_llm_generator(),
            text_sampler=self._get_text_sampler(),
            response_parser=QualityResponseParser(),
            clarity_coherence_prompt_template=read_text_resource(
                directory=PROMPTS_DIR, filename="clarity_coherence_prompt.txt"
            ),
            argumentation_conclusions_prompt_template=read_text_resource(
                directory=PROMPTS_DIR, filename="argumentation_conclusions_prompt.txt"
            ),
            resolver=self._get_quality_level_resolver(),
        )

    def _get_quality_level_resolver(self) -> QualityLevelResolver:
        return QualityLevelResolver(thresholds=self._get_quality_level_thresholds())

    def _get_quality_level_thresholds(self) -> QualityLevelThresholdsDTO:
        return QualityLevelThresholdsDTO(
            excellent_threshold=float(getenv("QUALITY_LEVEL_EXCELLENT_THRESHOLD", "9.0")),
            good_threshold=float(getenv("QUALITY_LEVEL_GOOD_THRESHOLD", "7.0")),
            acceptable_threshold=float(getenv("QUALITY_LEVEL_ACCEPTABLE_THRESHOLD", "5.0")),
            needs_improvement_threshold=float(
                getenv("QUALITY_LEVEL_NEEDS_IMPROVEMENT_THRESHOLD", "3.0")
            ),
        )

    def _get_text_sampler(self) -> QualityTextSampler:
        return QualityTextSampler(
            min_sample_word_count=int(getenv("QUALITY_MIN_SAMPLE_WORD_COUNT", "400")),
            text_sample_character_limit=int(getenv("QUALITY_TEXT_SAMPLE_CHARACTER_LIMIT", "8000")),
        )
