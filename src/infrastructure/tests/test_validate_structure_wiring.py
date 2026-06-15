from unittest import TestCase

from src.application.validate_structure_use_case import ValidateStructureUseCase
from src.infrastructure.wirings.validate_structure_wiring import ValidateStructureWiring


class TestValidateStructureWiring(TestCase):
    def setUp(self):
        self.wiring = ValidateStructureWiring()

    def test_create_use_case_returns_use_case_instance(self):
        use_case = self.wiring.create_use_case()
        self.assertIsInstance(use_case, ValidateStructureUseCase)

    def test_create_use_case_is_callable_multiple_times(self):
        use_case_1 = self.wiring.create_use_case()
        use_case_2 = self.wiring.create_use_case()
        self.assertIsInstance(use_case_1, ValidateStructureUseCase)
        self.assertIsInstance(use_case_2, ValidateStructureUseCase)
