from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.dimension_score_dto import DimensionScoreDTO
from src.domain.dtos.parsed_response_dto import ParsedResponseDTO
from src.domain.enums.quality_dimension import QualityDimension


class TestParsedResponseDTO(TestCase):
    def test_is_frozen_dataclass_extending_base_dto(self):
        self.assertTrue(issubclass(ParsedResponseDTO, BaseDTO))
        parsed_response = ParsedResponseDTO()
        with self.assertRaises(FrozenInstanceError):
            parsed_response.scores = {}

    def test_default_scores_and_matched_dimensions_are_empty(self):
        parsed_response = ParsedResponseDTO()
        self.assertEqual(parsed_response.scores, {})
        self.assertEqual(parsed_response.matched_dimensions, frozenset())

    def test_holds_scores_dict_and_matched_dimensions_frozenset(self):
        scores = {QualityDimension.CLARITY: DimensionScoreDTO(score=8.0, feedback="text")}
        matched_dimensions = frozenset({QualityDimension.CLARITY})
        parsed_response = ParsedResponseDTO(scores=scores, matched_dimensions=matched_dimensions)
        self.assertEqual(parsed_response.scores, scores)
        self.assertEqual(parsed_response.matched_dimensions, matched_dimensions)
