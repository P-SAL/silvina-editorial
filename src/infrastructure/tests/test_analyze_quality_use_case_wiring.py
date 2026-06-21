from typing import get_type_hints
from unittest import TestCase

from src.application.analyze_quality_use_case import AnalyzeQualityUseCase
from src.domain.ports.llm_generator_port import LlmGeneratorPort
from src.infrastructure.wirings.analyze_quality_use_case_wiring import (
    AnalyzeQualityUseCaseWiring,
)


class TestAnalyzeQualityUseCaseWiring(TestCase):
    def setUp(self):
        self.wiring = AnalyzeQualityUseCaseWiring()

    def test_create_use_case_returns_correct_type(self):
        use_case = self.wiring.create_use_case()
        self.assertIsInstance(use_case, AnalyzeQualityUseCase)

    def test_llm_generator_accessor_returns_port_type(self):
        hints = get_type_hints(self.wiring._get_llm_generator)
        self.assertEqual(hints["return"], LlmGeneratorPort)
