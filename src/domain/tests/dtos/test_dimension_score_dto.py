from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.dimension_score_dto import DimensionScoreDTO


class TestDimensionScoreDTO(TestCase):
    def test_is_frozen_dataclass_extending_base_dto(self):
        self.assertTrue(issubclass(DimensionScoreDTO, BaseDTO))
        dimension_score = DimensionScoreDTO(score=8.0, feedback="text")
        with self.assertRaises(FrozenInstanceError):
            dimension_score.score = 5.0

    def test_holds_score_and_feedback_fields(self):
        dimension_score = DimensionScoreDTO(score=8.0, feedback="text")
        self.assertEqual(dimension_score.score, 8.0)
        self.assertEqual(dimension_score.feedback, "text")
