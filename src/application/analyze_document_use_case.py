from src.application.analyze_quality_use_case import AnalyzeQualityUseCase
from src.application.check_grammar_use_case import CheckGrammarUseCase
from src.application.classify_article_use_case import ClassifyArticleUseCase
from src.application.extract_citations_use_case import ExtractCitationsUseCase
from src.application.extract_content_use_case import ExtractContentUseCase
from src.application.match_citations_use_case import MatchCitationsUseCase
from src.application.read_document_use_case import ReadDocumentUseCase
from src.application.validate_apa_use_case import ValidateApaUseCase
from src.application.validate_structure_use_case import ValidateStructureUseCase
from src.application.verify_eumic_use_case import VerifyEumicUseCase
from src.domain.dtos.report_input_dto import ReportInputDTO
from src.domain.enums.citation_type import CitationType
from src.domain.enums.section_name import SectionName
from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler
from src.domain.recommendation.recommendation_builder import RecommendationBuilder


class AnalyzeDocumentUseCase:
    """Orchestrator coordinating all document analysis use cases."""

    def __init__(
        self,
        read_document_use_case: ReadDocumentUseCase,
        extract_content_use_case: ExtractContentUseCase,
        extract_citations_use_case: ExtractCitationsUseCase,
        validate_apa_use_case: ValidateApaUseCase,
        check_grammar_use_case: CheckGrammarUseCase,
        classify_article_use_case: ClassifyArticleUseCase,
        analyze_quality_use_case: AnalyzeQualityUseCase,
        validate_structure_use_case: ValidateStructureUseCase,
        match_citations_use_case: MatchCitationsUseCase,
        verify_eumic_use_case: VerifyEumicUseCase,
        recommendation_builder: RecommendationBuilder,
    ) -> None:
        self._read_document_use_case = read_document_use_case
        self._extract_content_use_case = extract_content_use_case
        self._extract_citations_use_case = extract_citations_use_case
        self._validate_apa_use_case = validate_apa_use_case
        self._check_grammar_use_case = check_grammar_use_case
        self._classify_article_use_case = classify_article_use_case
        self._analyze_quality_use_case = analyze_quality_use_case
        self._validate_structure_use_case = validate_structure_use_case
        self._match_citations_use_case = match_citations_use_case
        self._verify_eumic_use_case = verify_eumic_use_case
        self._recommendation_builder = recommendation_builder

    @generic_error_handler
    def execute(self, document_path: str) -> ReportInputDTO:
        """Run the complete document analysis pipeline and return aggregated results."""
        paragraphs = self._read_document_use_case.execute(path=document_path)
        document_content = self._extract_content_use_case.execute(
            paragraphs=paragraphs, docx_path=document_path
        )
        citation_extraction = self._extract_citations_use_case.execute(docx_path=document_path)

        author_year_citations = [
            (c.text, c.location, paragraphs[c.location])
            for c in citation_extraction.citations
            if c.citation_type == CitationType.AUTHOR_YEAR
        ]
        apa_validation = self._validate_apa_use_case.execute(citations=author_year_citations)

        grammar = self._check_grammar_use_case.execute(paragraphs=paragraphs)
        classification = self._classify_article_use_case.execute(document_content=document_content)
        quality = self._analyze_quality_use_case.execute(
            document_content=document_content, article_type=classification.article_type
        )

        effective_type = classification.effective_structure_type
        has_references = len(citation_extraction.references) > 0
        structure = self._validate_structure_use_case.execute(
            document_content=document_content,
            article_type=effective_type,
            has_references=has_references,
        )

        try:
            section_name = SectionName(citation_extraction.section_type)
        except ValueError:
            section_name = SectionName.REFERENCES

        matched_citations = self._match_citations_use_case.execute(
            citations=citation_extraction.citations,
            references=citation_extraction.references,
            section_type=section_name,
        )

        eumic_violations = self._verify_eumic_use_case.execute(
            docx_path=document_path,
            word_count=document_content.word_count,
        )

        recommendations, verdict = self._recommendation_builder.build(
            classification=classification,
            quality=quality,
            structure=structure,
            citations=matched_citations,
            apa_validation=apa_validation,
            grammar=grammar,
        )

        return ReportInputDTO(
            filename=document_path,
            document_content=document_content,
            classification=classification,
            quality=quality,
            grammar=grammar,
            structure=structure,
            citations=matched_citations,
            apa_validation=apa_validation,
            recommendations=recommendations,
            verdict=verdict,
            eumic_violations=eumic_violations,
        )
