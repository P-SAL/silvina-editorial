from unittest import TestCase
from unittest.mock import MagicMock

from src.application.analyze_document_use_case import AnalyzeDocumentUseCase
from src.domain.dtos.citation_dto import CitationDTO
from src.domain.dtos.eumic_violation_dto import EumicViolationDTO
from src.domain.dtos.report_input_dto import ReportInputDTO
from src.domain.enums.article_type import ArticleType
from src.domain.enums.citation_type import CitationType


def _make_citation(text="(Smith, 2020)", citation_type=CitationType.AUTHOR_YEAR, location=0):
    return CitationDTO(text=text, citation_type=citation_type, location=location)


def _make_classification(article_type=ArticleType.POPULAR_SCIENCE, reasoning="Test"):
    m = MagicMock()
    m.article_type = article_type
    m.effective_structure_type = article_type
    m.reasoning = reasoning
    return m


def _make_content(word_count=2500, paragraphs=None):
    m = MagicMock()
    m.word_count = word_count
    m.paragraphs = paragraphs if paragraphs is not None else ["Para0", "Para1"]
    return m


class TestAnalyzeDocumentUseCase(TestCase):
    def _make_use_case(self, **overrides):
        document_text_port = MagicMock()
        content_extraction_port = MagicMock()
        character_count_port = MagicMock()
        citation_extraction_port = MagicMock()
        reference_extraction_port = MagicMock()
        grammar_check_port = MagicMock()
        document_format_inspection_port = MagicMock()
        apa_validator = MagicMock()
        article_classifier = MagicMock()
        quality_analyzer = MagicMock()
        structure_validator = MagicMock()
        citation_matcher = MagicMock()
        recommendation_builder = MagicMock()

        document_text_port.read_paragraphs.return_value = ["Paragraph 0", "Paragraph 1"]
        content_extraction_port.extract.return_value = _make_content()
        character_count_port.count.return_value = None
        citation_extraction_port.extract_citations.return_value = []
        reference_extraction_port.extract_references.return_value = ([], "Referencias")
        apa_validator.validate_all_citations.return_value = []
        grammar_check_port.check.return_value = []
        article_classifier.classify.return_value = _make_classification()
        quality_analyzer.analyze.return_value = MagicMock(overall_score=8.0)
        structure_validator.validate.return_value = ([], [])
        citation_matcher.match_citations_to_references.return_value = MagicMock(
            total_citations=5, matched_count=5, unmatched_count=0
        )
        document_format_inspection_port.inspect.return_value = []
        recommendation_builder.build.return_value = ([], MagicMock())

        mocks = {
            "document_text_port": document_text_port,
            "content_extraction_port": content_extraction_port,
            "character_count_port": character_count_port,
            "citation_extraction_port": citation_extraction_port,
            "reference_extraction_port": reference_extraction_port,
            "grammar_check_port": grammar_check_port,
            "document_format_inspection_port": document_format_inspection_port,
            "apa_validator": apa_validator,
            "article_classifier": article_classifier,
            "quality_analyzer": quality_analyzer,
            "structure_validator": structure_validator,
            "citation_matcher": citation_matcher,
            "recommendation_builder": recommendation_builder,
        }
        mocks.update(overrides)
        use_case = AnalyzeDocumentUseCase(**mocks)
        return use_case, mocks

    def test_execute_returns_report_input_dto(self):
        use_case, _ = self._make_use_case()
        result = use_case.execute(document_path="test.docx")
        self.assertIsInstance(result, ReportInputDTO)

    def test_execute_calls_all_ports_and_services_once(self):
        use_case, mocks = self._make_use_case()
        use_case.execute(document_path="test.docx")

        mocks["document_text_port"].read_paragraphs.assert_called_once_with(path="test.docx")
        mocks["content_extraction_port"].extract.assert_called_once()
        mocks["character_count_port"].count.assert_called_once_with(docx_path="test.docx")
        mocks["citation_extraction_port"].extract_citations.assert_called_once_with(
            docx_path="test.docx"
        )
        mocks["reference_extraction_port"].extract_references.assert_called_once_with(
            docx_path="test.docx"
        )
        mocks["apa_validator"].validate_all_citations.assert_not_called()
        mocks["grammar_check_port"].check.assert_called_once()
        mocks["article_classifier"].classify.assert_called_once()
        mocks["quality_analyzer"].analyze.assert_called_once()
        mocks["structure_validator"].validate.assert_called_once()
        mocks["citation_matcher"].match_citations_to_references.assert_called_once()
        mocks["document_format_inspection_port"].inspect.assert_called_once()
        mocks["recommendation_builder"].build.assert_called_once()

    def test_only_author_year_citations_sent_to_apa_validator(self):
        author_year = _make_citation(
            text="(Smith, 2020)", citation_type=CitationType.AUTHOR_YEAR, location=0
        )
        numeric = _make_citation(text="[1]", citation_type=CitationType.NUMERIC, location=1)

        citation_extraction_port = MagicMock()
        citation_extraction_port.extract_citations.return_value = [author_year, numeric]

        use_case, mocks = self._make_use_case(citation_extraction_port=citation_extraction_port)
        use_case.execute(document_path="test.docx")

        sent_citations = mocks["apa_validator"].validate_all_citations.call_args.kwargs["citations"]
        self.assertEqual(len(sent_citations), 1)
        self.assertEqual(sent_citations[0][0], "(Smith, 2020)")

    def test_citation_tuple_includes_paragraph_text_at_location(self):
        paragraphs = ["Para 0", "Para 1", "Para 2"]
        document_text_port = MagicMock()
        document_text_port.read_paragraphs.return_value = paragraphs

        citation = _make_citation(
            text="(Jones, 2019)", citation_type=CitationType.AUTHOR_YEAR, location=2
        )
        citation_extraction_port = MagicMock()
        citation_extraction_port.extract_citations.return_value = [citation]

        use_case, mocks = self._make_use_case(
            document_text_port=document_text_port,
            citation_extraction_port=citation_extraction_port,
        )
        use_case.execute(document_path="test.docx")

        sent_citations = mocks["apa_validator"].validate_all_citations.call_args.kwargs["citations"]
        self.assertEqual(sent_citations[0], ("(Jones, 2019)", 2, "Para 2"))

    def test_structure_validated_with_effective_structure_type(self):
        classification = _make_classification(
            article_type=ArticleType.SCIENTIFIC,
        )
        classification.effective_structure_type = ArticleType.POPULAR_SCIENCE

        article_classifier = MagicMock()
        article_classifier.classify.return_value = classification

        use_case, mocks = self._make_use_case(article_classifier=article_classifier)
        use_case.execute(document_path="test.docx")

        validate_call = mocks["structure_validator"].validate.call_args
        passed_type = validate_call.kwargs["article_type"]
        self.assertEqual(passed_type, ArticleType.POPULAR_SCIENCE)

    def test_eumic_violations_included_in_report_input_dto(self):
        violation = MagicMock(spec=EumicViolationDTO)
        document_format_inspection_port = MagicMock()
        document_format_inspection_port.inspect.return_value = [violation]

        use_case, _ = self._make_use_case(
            document_format_inspection_port=document_format_inspection_port
        )
        result = use_case.execute(document_path="test.docx")

        self.assertEqual(len(result.eumic_violations), 1)
        self.assertIs(result.eumic_violations[0], violation)

    def test_eumic_violations_do_not_halt_execution(self):
        violation = MagicMock(spec=EumicViolationDTO)
        document_format_inspection_port = MagicMock()
        document_format_inspection_port.inspect.return_value = [violation]

        use_case, mocks = self._make_use_case(
            document_format_inspection_port=document_format_inspection_port
        )
        result = use_case.execute(document_path="test.docx")

        self.assertIsInstance(result, ReportInputDTO)
        mocks["recommendation_builder"].build.assert_called_once()

    def test_report_contains_recommendations_from_builder(self):
        use_case, _ = self._make_use_case()
        result = use_case.execute(document_path="test.docx")
        self.assertEqual(result.recommendations, [])

    def test_has_references_is_true_when_references_present(self):
        from src.domain.enums.section_name import SectionName

        reference_extraction_port = MagicMock()
        reference_extraction_port.extract_references.return_value = ([MagicMock()], "Referencias")

        structure_validator = MagicMock()
        structure_validator.validate.return_value = ([], [SectionName.REFERENCES])

        use_case, _ = self._make_use_case(
            reference_extraction_port=reference_extraction_port,
            structure_validator=structure_validator,
        )
        result = use_case.execute(document_path="test.docx")

        self.assertNotIn(SectionName.REFERENCES, result.structure.missing_sections)

    def test_has_references_is_false_when_no_references(self):
        from src.domain.enums.section_name import SectionName

        reference_extraction_port = MagicMock()
        reference_extraction_port.extract_references.return_value = ([], "Referencias")

        structure_validator = MagicMock()
        structure_validator.validate.return_value = ([], [SectionName.REFERENCES])

        use_case, _ = self._make_use_case(
            reference_extraction_port=reference_extraction_port,
            structure_validator=structure_validator,
        )
        result = use_case.execute(document_path="test.docx")

        self.assertIn(SectionName.REFERENCES, result.structure.missing_sections)

    def test_apa_validation_skipped_when_no_author_year_citations(self):
        use_case, mocks = self._make_use_case()
        result = use_case.execute(document_path="test.docx")

        mocks["apa_validator"].validate_all_citations.assert_not_called()
        self.assertTrue(result.apa_validation.is_valid)
        self.assertEqual(result.apa_validation.violation_count, 0)

    def test_section_type_defaults_to_references_when_invalid(self):
        reference_extraction_port = MagicMock()
        reference_extraction_port.extract_references.return_value = ([], "Invalid Section")

        use_case, mocks = self._make_use_case(reference_extraction_port=reference_extraction_port)
        use_case.execute(document_path="test.docx")

        from src.domain.enums.section_name import SectionName

        match_call = mocks["citation_matcher"].match_citations_to_references.call_args
        section_arg = match_call.kwargs["section_type"]
        self.assertEqual(section_arg, SectionName.REFERENCES)

    def test_character_count_unavailable_falls_back_to_base_content(self):
        from src.domain.exceptions.count_errors import CharacterCountUnavailable

        character_count_port = MagicMock()
        character_count_port.count.side_effect = CharacterCountUnavailable

        use_case, mocks = self._make_use_case(character_count_port=character_count_port)
        result = use_case.execute(document_path="test.docx")

        self.assertIsInstance(result, ReportInputDTO)
