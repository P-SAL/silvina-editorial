from unittest import TestCase

from src.application.verify_eumic_use_case import VerifyEumicUseCase
from src.infrastructure.adapters.document.docx_eumic_adapter import DocxEumicAdapter
from src.infrastructure.wirings.verify_eumic_use_case_wiring import VerifyEumicUseCaseWiring


class TestVerifyEumicUseCaseWiring(TestCase):
    def test_create_use_case_returns_verify_eumic_use_case_instance(self):
        result = VerifyEumicUseCaseWiring().create_use_case()
        self.assertIsInstance(result, VerifyEumicUseCase)

    def test_create_use_case_wires_docx_eumic_adapter_as_format_inspection_port(self):
        result = VerifyEumicUseCaseWiring().create_use_case()
        self.assertIsInstance(result._format_inspection_port, DocxEumicAdapter)
