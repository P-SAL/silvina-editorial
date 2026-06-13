from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.citation_analysis_result_dto import CitationAnalysisResult


class TestCitationAnalysisResult(TestCase):
    def test_citation_analysis_result_is_subclass_of_base_dto(self):
        self.assertTrue(issubclass(CitationAnalysisResult, BaseDTO))

    def test_citation_analysis_result_is_immutable(self):
        result = CitationAnalysisResult(
            total_citations=5,
            total_references=4,
            matched_count=3,
            unmatched_count=2,
        )
        with self.assertRaises(FrozenInstanceError):
            result.total_citations = 0

    def test_str_with_citations(self):
        result = CitationAnalysisResult(
            total_citations=10,
            total_references=8,
            matched_count=8,
            unmatched_count=2,
        )
        self.assertEqual(str(result), "Citations: 10 (80.0% matched)")

    def test_str_with_zero_citations(self):
        result = CitationAnalysisResult(
            total_citations=0,
            total_references=0,
            matched_count=0,
            unmatched_count=0,
        )
        self.assertEqual(str(result), "Citations: 0 (0.0% matched)")
