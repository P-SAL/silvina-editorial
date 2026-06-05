"""
Unit tests for eumic_verifier.py
Uses mock python-docx Document objects to avoid file I/O.
"""
import sys
import os
import unittest
from unittest.mock import MagicMock, patch, PropertyMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

# Defensive guard for COM mocks
if 'win32com' not in sys.modules:
    _wc = MagicMock(); _wcc = MagicMock(); _wc.client = _wcc
    sys.modules.update({'win32com': _wc, 'win32com.client': _wcc, 'pythoncom': MagicMock()})

from eumic_verifier import EumicVerifier, EumicViolation, EumicSeverity, verify_eumic_compliance


def _make_mock_doc(paragraphs_text=None, tables_count=0):
    """Build a minimal mock python-docx Document."""
    doc = MagicMock()

    # Sections — use 2.5 cm margins (compliant)
    from docx.shared import Cm
    section = MagicMock()
    section.top_margin.twips = Cm(2.5).twips
    section.bottom_margin.twips = Cm(2.5).twips
    section.left_margin.twips = Cm(2.5).twips
    section.right_margin.twips = Cm(2.5).twips
    doc.sections = [section]

    # Paragraphs
    paras = []
    for text in (paragraphs_text or []):
        para = MagicMock()
        para.text = text
        para.runs = []
        para.alignment = None  # Not justified — we won't check alignment for most tests
        paras.append(para)
    doc.paragraphs = paras

    # Tables
    doc.tables = [MagicMock() for _ in range(tables_count)]

    # No images
    doc.part.rels = {}

    return doc


def _make_mock_content(word_count=3000, has_abstract=True, has_keywords=True):
    """Build a minimal mock DocumentContent."""
    content = MagicMock()
    content.word_count = word_count
    return content


class TestEumicVerifierCompliant(unittest.TestCase):
    """A document with all required sections should produce no CRITICAL violations."""

    def setUp(self):
        self.verifier = EumicVerifier()

    def test_compliant_document_no_critical_violations(self):
        paras = [
            'Resumen',
            'Este es el resumen del artículo de investigación.',
            'Palabras clave: investigación, académico, ciencia',
        ]
        doc = _make_mock_doc(paras)
        content = _make_mock_content()
        violations = self.verifier.verify_document(doc, content)
        critical = [v for v in violations if v.severity == EumicSeverity.CRITICAL]
        self.assertEqual(critical, [])

    def test_verify_document_returns_list(self):
        doc = _make_mock_doc(['Resumen', 'Texto del resumen.', 'Palabras clave: a, b, c'])
        content = _make_mock_content()
        result = self.verifier.verify_document(doc, content)
        self.assertIsInstance(result, list)


class TestEumicVerifierMissingAbstract(unittest.TestCase):
    """Missing abstract in a long document should produce CRITICAL violation."""

    def setUp(self):
        self.verifier = EumicVerifier()

    def test_missing_abstract_flagged_as_critical(self):
        paras = ['Palabras clave: a, b, c', 'Introducción', 'Texto de introducción.']
        doc = _make_mock_doc(paras)
        content = _make_mock_content(word_count=2000)
        violations = self.verifier.verify_document(doc, content)
        msgs = [v.message for v in violations if v.severity == EumicSeverity.CRITICAL]
        self.assertTrue(any('Resumen' in m or 'Abstract' in m for m in msgs))

    def test_short_document_no_abstract_not_critical(self):
        """Short docs (< 1000 words) should not require abstract."""
        paras = ['Palabras clave: a, b, c']
        doc = _make_mock_doc(paras)
        content = _make_mock_content(word_count=500)
        violations = self.verifier.verify_document(doc, content)
        critical = [v for v in violations if v.severity == EumicSeverity.CRITICAL]
        # Should not flag CRITICAL for short docs
        abstract_critical = [v for v in critical if 'Resumen' in v.message or 'Abstract' in v.message]
        self.assertEqual(abstract_critical, [])


class TestEumicVerifierMissingKeywords(unittest.TestCase):
    """Missing keywords in a long document should be CRITICAL."""

    def setUp(self):
        self.verifier = EumicVerifier()

    def test_missing_keywords_flagged_as_critical(self):
        paras = ['Resumen', 'Texto del resumen del artículo académico.']
        doc = _make_mock_doc(paras)
        content = _make_mock_content(word_count=2000)
        violations = self.verifier.verify_document(doc, content)
        msgs = [v.message for v in violations if v.severity == EumicSeverity.CRITICAL]
        self.assertTrue(any('clave' in m.lower() or 'keyword' in m.lower() for m in msgs))


class TestEumicVerifierFormatReport(unittest.TestCase):
    """format_violations_report returns empty string for no violations."""

    def setUp(self):
        self.verifier = EumicVerifier()

    def test_empty_report_for_no_violations(self):
        report = self.verifier.format_violations_report([])
        self.assertEqual(report, '')

    def test_report_contains_critical_section(self):
        violation = EumicViolation(
            category='Test',
            message='Test violation',
            severity=EumicSeverity.CRITICAL,
            details='Details here'
        )
        report = self.verifier.format_violations_report([violation])
        self.assertIn('CRÍTICO', report)

    def test_convenience_function_returns_string(self):
        doc = _make_mock_doc(['Resumen', 'Texto.', 'Palabras clave: a, b, c'])
        content = _make_mock_content()
        result = verify_eumic_compliance(doc, content)
        self.assertIsInstance(result, str)


if __name__ == '__main__':
    unittest.main()
