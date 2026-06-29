from unittest import TestCase

from src.application.verify_eumic_use_case import VerifyEumicUseCase
from src.domain.dtos.eumic_violation_dto import EumicViolationDTO
from src.domain.enums.severity_level import SeverityLevel
from src.domain.exceptions.base_src_error import SrcGenericError
from src.domain.tests.document.fake_document_format_inspection_port import (
    FakeDocumentFormatInspectionPort,
)


def _make_violation(message: str = "test violation") -> EumicViolationDTO:
    return EumicViolationDTO(
        category="Formato General",
        message=message,
        severity=SeverityLevel.WARNING,
    )


class TestVerifyEumicUseCase(TestCase):
    def test_execute_returns_empty_list_when_port_returns_no_violations(self):
        use_case = VerifyEumicUseCase(
            format_inspection_port=FakeDocumentFormatInspectionPort(violations=[])
        )
        result = use_case.execute(docx_path="doc.docx", word_count=2000)
        self.assertEqual(result, [])

    def test_execute_returns_violations_from_port(self):
        violations = [_make_violation("Margen incorrecto"), _make_violation("Fuente inválida")]
        use_case = VerifyEumicUseCase(
            format_inspection_port=FakeDocumentFormatInspectionPort(violations=violations)
        )
        result = use_case.execute(docx_path="doc.docx", word_count=2000)
        self.assertEqual(result, violations)

    def test_execute_passes_docx_path_to_port(self):
        captured = {}

        class CapturingPort(FakeDocumentFormatInspectionPort):
            def inspect(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
                captured["docx_path"] = docx_path
                captured["word_count"] = word_count
                return []

        use_case = VerifyEumicUseCase(format_inspection_port=CapturingPort())
        use_case.execute(docx_path="/path/to/doc.docx", word_count=1500)
        self.assertEqual(captured["docx_path"], "/path/to/doc.docx")
        self.assertEqual(captured["word_count"], 1500)

    def test_execute_propagates_port_exception_as_src_generic_error(self):
        use_case = VerifyEumicUseCase(
            format_inspection_port=FakeDocumentFormatInspectionPort(
                error=RuntimeError("unexpected")
            )
        )
        with self.assertRaises(SrcGenericError):
            use_case.execute(docx_path="doc.docx", word_count=2000)
