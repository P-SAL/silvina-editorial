import os
from unittest import TestCase
from unittest.mock import patch

from src.infrastructure.adapters.report.docx_report_settings import DocxReportSettings


class TestDocxReportSettings(TestCase):
    def test_default_words_per_page_is_250(self):
        env_without = {k: v for k, v in os.environ.items() if k != "REPORT_WORDS_PER_PAGE"}
        with patch.dict(os.environ, env_without, clear=True):
            settings = DocxReportSettings()
        self.assertEqual(settings.words_per_page, 250)

    def test_default_max_errors_displayed_is_5(self):
        env_without = {k: v for k, v in os.environ.items() if k != "REPORT_MAX_ERRORS_DISPLAYED"}
        with patch.dict(os.environ, env_without, clear=True):
            settings = DocxReportSettings()
        self.assertEqual(settings.max_errors_displayed, 5)

    def test_default_context_truncation_limit_is_150(self):
        env_without = {
            k: v for k, v in os.environ.items() if k != "REPORT_CONTEXT_TRUNCATION_LIMIT"
        }
        with patch.dict(os.environ, env_without, clear=True):
            settings = DocxReportSettings()
        self.assertEqual(settings.context_truncation_limit, 150)

    def test_default_max_replacements_is_3(self):
        env_without = {k: v for k, v in os.environ.items() if k != "REPORT_MAX_REPLACEMENTS"}
        with patch.dict(os.environ, env_without, clear=True):
            settings = DocxReportSettings()
        self.assertEqual(settings.max_replacements, 3)

    def test_env_var_overrides_words_per_page(self):
        with patch.dict(os.environ, {"REPORT_WORDS_PER_PAGE": "100"}):
            settings = DocxReportSettings()
        self.assertEqual(settings.words_per_page, 100)

    def test_env_var_overrides_max_errors_displayed(self):
        with patch.dict(os.environ, {"REPORT_MAX_ERRORS_DISPLAYED": "2"}):
            settings = DocxReportSettings()
        self.assertEqual(settings.max_errors_displayed, 2)

    def test_env_var_overrides_context_truncation_limit(self):
        with patch.dict(os.environ, {"REPORT_CONTEXT_TRUNCATION_LIMIT": "10"}):
            settings = DocxReportSettings()
        self.assertEqual(settings.context_truncation_limit, 10)

    def test_env_var_overrides_max_replacements(self):
        with patch.dict(os.environ, {"REPORT_MAX_REPLACEMENTS": "1"}):
            settings = DocxReportSettings()
        self.assertEqual(settings.max_replacements, 1)
