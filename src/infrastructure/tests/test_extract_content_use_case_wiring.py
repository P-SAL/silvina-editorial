from inspect import getsource
from unittest import TestCase

from src.application.extract_content_use_case import ExtractContentUseCase
from src.infrastructure.adapters.document.paragraph_content_adapter import ParagraphContentAdapter
from src.infrastructure.adapters.document.win32com_word_count_adapter import (
    Win32ComWordCountAdapter,
)
from src.infrastructure.wirings.extract_content_use_case_wiring import ExtractContentUseCaseWiring


class TestExtractContentUseCaseWiring(TestCase):
    def test_create_use_case_returns_extract_content_use_case_with_correct_adapters(self):
        use_case = ExtractContentUseCaseWiring().create_use_case()

        self.assertIsInstance(use_case, ExtractContentUseCase)
        self.assertIsInstance(use_case._extraction_port, ParagraphContentAdapter)
        self.assertIsInstance(use_case._count_port, Win32ComWordCountAdapter)

    def test_adapter_logic_confined_to_private_methods(self):
        source = getsource(ExtractContentUseCaseWiring.create_use_case)

        self.assertNotIn("ParagraphContent", source)
        self.assertNotIn("Win32Com", source)
