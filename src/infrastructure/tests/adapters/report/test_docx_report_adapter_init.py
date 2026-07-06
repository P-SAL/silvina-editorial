from unittest import TestCase
from unittest.mock import patch

from src.domain.exceptions.report_errors import ReportExportUnavailable
from src.infrastructure.adapters.report.docx_report_adapter import DocxReportAdapter


class TestDocxReportAdapterInit(TestCase):
    def test_init_succeeds_without_logo_path(self):
        adapter = DocxReportAdapter(logo_path=None)
        self.assertIsInstance(adapter, DocxReportAdapter)

    @patch("src.infrastructure.adapters.report.docx_report_adapter.DOCX_AVAILABLE", False)
    def test_init_raises_report_export_unavailable_when_docx_missing(self):
        with self.assertRaises(ReportExportUnavailable):
            DocxReportAdapter(logo_path=None)
