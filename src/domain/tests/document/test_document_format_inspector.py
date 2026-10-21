from unittest import TestCase

from src.domain.document.document_format_inspector import DocumentFormatInspector
from src.domain.dtos.eumic_violation_dto import EumicViolationDTO
from src.domain.enums.severity_level import SeverityLevel
from src.domain.tests.document.fake_document_format_inspection_port import (
    FakeDocumentFormatInspectionPort,
)


class TestDocumentFormatInspector(TestCase):
    def test_inspect_returns_violations_from_port(self):
        violation = EumicViolationDTO(
            category="font", message="Wrong font", severity=SeverityLevel.WARNING
        )
        port = FakeDocumentFormatInspectionPort(violations=[violation])
        inspector = DocumentFormatInspector(document_format_inspection_port=port)

        result = inspector.inspect(docx_path="test.docx", word_count=1000)

        self.assertEqual(result, [violation])

    def test_inspect_returns_empty_list_when_no_violations(self):
        port = FakeDocumentFormatInspectionPort(violations=[])
        inspector = DocumentFormatInspector(document_format_inspection_port=port)

        result = inspector.inspect(docx_path="test.docx", word_count=500)

        self.assertEqual(result, [])
