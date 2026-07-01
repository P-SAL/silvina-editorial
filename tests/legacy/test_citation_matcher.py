"""
Unit tests for business_logic/citation_matcher.py
Uses in-memory Citation and Reference lists; no file I/O.
"""

import sys
import os
import unittest
from unittest.mock import MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

# Defensive guard for COM mocks
if "win32com" not in sys.modules:
    _wc = MagicMock()
    _wcc = MagicMock()
    _wc.client = _wcc
    sys.modules.update({"win32com": _wc, "win32com.client": _wcc, "pythoncom": MagicMock()})

from business_logic.citation_matcher import CitationMatcher
from domain.models import Citation, Reference
from domain.enums import CitationType


def _citation(author, year="2020", text=None):
    return Citation(
        text=text or f"({author}, {year})",
        citation_type=CitationType.AUTHOR_YEAR,
        location=0,
        author=author,
        year=year,
    )


def _reference(text):
    return Reference(text=text)


class TestCitationMatcherOrphans(unittest.TestCase):
    def test_all_matched_no_orphans(self):
        citations = [_citation("García")]
        references = [_reference("García, J. (2020). Título. Revista, 1, 1-10.")]
        matcher = CitationMatcher(citations, references)
        self.assertEqual(matcher.find_orphaned_citations(), [])

    def test_orphaned_citation_detected(self):
        citations = [_citation("Pérez")]
        references = [_reference("García, J. (2020). Título. Revista, 1, 1-10.")]
        matcher = CitationMatcher(citations, references)
        orphaned = matcher.find_orphaned_citations()
        self.assertEqual(len(orphaned), 1)
        self.assertEqual(orphaned[0].author, "Pérez")

    def test_orphaned_reference_detected(self):
        citations = [_citation("García")]
        references = [
            _reference("García, J. (2020). Título. Revista, 1, 1-10."),
            _reference("López, M. (2021). Otro título. Journal, 2, 5-15."),
        ]
        matcher = CitationMatcher(citations, references)
        orphaned_refs = matcher.find_orphaned_references()
        # López is in references but not cited
        self.assertEqual(len(orphaned_refs), 1)
        self.assertIn("López", orphaned_refs[0].text)

    def test_no_orphans_when_all_match(self):
        citations = [_citation("García"), _citation("López")]
        references = [
            _reference("García, J. (2020). Título. Revista, 1, 1-10."),
            _reference("López, M. (2021). Otro. Journal, 2, 5-15."),
        ]
        matcher = CitationMatcher(citations, references)
        self.assertEqual(matcher.find_orphaned_citations(), [])
        self.assertEqual(matcher.find_orphaned_references(), [])


class TestCitationMatcherFootnoteSkip(unittest.TestCase):
    """Footnotes should be skipped in citation matching."""

    def test_footnotes_not_counted_as_orphans(self):
        footnote = Citation(
            text="[Footnote 1]",
            citation_type=CitationType.FOOTNOTE,
            location=0,
            author=None,
            year=None,
        )
        references = [_reference("García, J. (2020). Título. Revista.")]
        matcher = CitationMatcher([footnote], references)
        self.assertEqual(matcher.find_orphaned_citations(), [])


class TestCitationMatcherNormalization(unittest.TestCase):
    """Author normalization should handle et al. and y/and."""

    def test_et_al_matches_first_author(self):
        citations = [_citation("García et al.")]
        references = [_reference("García, J. (2020). Título. Revista.")]
        matcher = CitationMatcher(citations, references)
        self.assertEqual(matcher.find_orphaned_citations(), [])

    def test_non_author_pattern_skipped(self):
        non_author = _citation("UNESCO 2020", text="(UNESCO 2020)")
        non_author.author = "UNESCO 2020"
        references = []
        matcher = CitationMatcher([non_author], references)
        # non-author patterns should not appear as orphans
        orphaned = matcher.find_orphaned_citations()
        self.assertEqual(orphaned, [])


class TestCitationMatcherReport(unittest.TestCase):
    def test_generate_report_returns_string(self):
        citations = [_citation("García")]
        references = [_reference("García, J. (2020). Título. Revista.")]
        matcher = CitationMatcher(citations, references)
        report = matcher.generate_report("Referencias")
        self.assertIsInstance(report, str)
        self.assertIn("INTEGRIDAD", report)

    def test_match_citations_to_references_returns_result(self):
        citations = [_citation("García")]
        references = [_reference("García, J. (2020). Título. Revista.")]
        matcher = CitationMatcher(citations, references)
        result = matcher.match_citations_to_references("Referencias")
        self.assertEqual(result.total_citations, 1)
        self.assertEqual(result.total_references, 1)

    def test_empty_inputs_no_crash(self):
        matcher = CitationMatcher([], [])
        self.assertEqual(matcher.find_orphaned_citations(), [])
        self.assertEqual(matcher.find_orphaned_references(), [])

    def test_generate_report_with_orphaned_citations(self):
        citations = [_citation("Pérez")]
        references = [_reference("García, J. (2020). Introducción. Revista, 1, 1-10.")]
        matcher = CitationMatcher(citations, references)
        report = matcher.generate_report("Referencias")
        self.assertIsInstance(report, str)
        self.assertIn("critical", report.lower())


if __name__ == "__main__":
    unittest.main()
