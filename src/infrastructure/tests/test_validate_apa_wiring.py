from unittest import TestCase

from src.infrastructure.wirings.validate_apa_wiring import ValidateApaWiring
from src.application.validate_apa_use_case import ValidateApaUseCase
from src.domain.dtos.apa_validation_result_dto import ApaValidationResultDTO


class TestValidateApaWiring(TestCase):
    def setUp(self):
        self.wiring = ValidateApaWiring()

    def test_s15a_create_use_case_returns_correct_type(self):
        use_case = self.wiring.create_use_case()
        self.assertIsInstance(use_case, ValidateApaUseCase)

    def test_s15b_use_case_execute_returns_apa_validation_result(self):
        use_case = self.wiring.create_use_case()
        result = use_case.execute(citations=[])
        self.assertIsInstance(result, ApaValidationResultDTO)
