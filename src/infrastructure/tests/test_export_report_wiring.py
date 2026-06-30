from unittest import TestCase
from unittest.mock import patch

from src.application.export_report_use_case import ExportReportUseCase
from src.infrastructure.adapters.report.docx_report_adapter import DocxReportAdapter
from src.infrastructure.wirings.export_report_wiring import ExportReportWiring


class TestExportReportWiring(TestCase):
    def test_create_use_case_returns_export_report_use_case_instance(self):
        result = ExportReportWiring().create_use_case()
        self.assertIsInstance(result, ExportReportUseCase)

    def test_create_use_case_wires_docx_report_adapter_as_port(self):
        result = ExportReportWiring().create_use_case()
        self.assertIsInstance(result._report_export_port, DocxReportAdapter)

    @patch(
        "src.infrastructure.wirings.export_report_wiring.DocxReportAdapter",
        side_effect=ImportError("No module named 'docx'"),
    )
    def test_create_use_case_propagates_import_error_when_docx_unavailable(self, _):
        with self.assertRaises(ImportError):
            ExportReportWiring().create_use_case()
