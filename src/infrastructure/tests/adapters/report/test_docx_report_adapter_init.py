from unittest import TestCase
from unittest.mock import MagicMock, patch

from src.infrastructure.adapters.report.docx_report_adapter import DocxReportAdapter


class TestDocxReportAdapterInit(TestCase):
    def test_init_succeeds_without_logo_path(self):
        adapter = DocxReportAdapter(logo_path=None)
        self.assertIsInstance(adapter, DocxReportAdapter)

    @patch(
        "src.infrastructure.adapters.report.docx_report_adapter.Document",
        side_effect=ImportError("No module named 'docx'"),
    )
    def test_export_propagates_import_error_when_docx_unavailable(self, _):
        adapter = DocxReportAdapter(logo_path=None)
        with self.assertRaises(ImportError):
            adapter.export(report_input=MagicMock(), output_path="out.docx")
