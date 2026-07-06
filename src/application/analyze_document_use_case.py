from dataclasses import replace

from src.domain.citation.apa_validator import ApaValidator
from src.domain.citation.citation_matcher import CitationMatcher
from src.domain.classification.article_classifier import ArticleClassifier
from src.domain.document.character_count_port import CharacterCountPort
from src.domain.document.citation_extraction_port import CitationExtractionPort
from src.domain.document.content_extraction_port import ContentExtractionPort
from src.domain.document.document_format_inspection_port import DocumentFormatInspectionPort
from src.domain.document.document_text_port import DocumentTextPort
from src.domain.document.reference_extraction_port import ReferenceExtractionPort
from src.domain.dtos.apa_validation_result_dto import ApaValidationResultDTO
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.dtos.grammar_check_result_dto import GrammarCheckResultDTO
from src.domain.dtos.report_input_dto import ReportInputDTO
from src.domain.dtos.structure_validation_result_dto import StructureValidationResultDTO
from src.domain.enums.article_type import ArticleType
from src.domain.enums.citation_type import CitationType
from src.domain.enums.section_name import SectionName
from src.domain.exceptions.count_errors import CharacterCountUnavailable
from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler
from src.domain.exceptions.document_errors import DocumentEmpty
from src.domain.grammar.grammar_check_port import GrammarCheckPort
from src.domain.grammar.grammar_score_level import GrammarScoreLevel
from src.domain.quality.quality_analyzer import QualityAnalyzer
from src.domain.recommendation.recommendation_builder import RecommendationBuilder
from src.domain.structure.structure_validator import StructureValidator


class AnalyzeDocumentUseCase:
    """Orchestrates the complete academic document analysis pipeline."""

    def __init__(
        self,
        document_text_port: DocumentTextPort,
        content_extraction_port: ContentExtractionPort,
        character_count_port: CharacterCountPort,
        citation_extraction_port: CitationExtractionPort,
        reference_extraction_port: ReferenceExtractionPort,
        grammar_check_port: GrammarCheckPort,
        document_format_inspection_port: DocumentFormatInspectionPort,
        apa_validator: ApaValidator,
        article_classifier: ArticleClassifier,
        quality_analyzer: QualityAnalyzer,
        structure_validator: StructureValidator,
        citation_matcher: CitationMatcher,
        recommendation_builder: RecommendationBuilder,
    ) -> None:
        self._document_text_port = document_text_port
        self._content_extraction_port = content_extraction_port
        self._character_count_port = character_count_port
        self._citation_extraction_port = citation_extraction_port
        self._reference_extraction_port = reference_extraction_port
        self._grammar_check_port = grammar_check_port
        self._document_format_inspection_port = document_format_inspection_port
        self._apa_validator = apa_validator
        self._article_classifier = article_classifier
        self._quality_analyzer = quality_analyzer
        self._structure_validator = structure_validator
        self._citation_matcher = citation_matcher
        self._recommendation_builder = recommendation_builder

    @generic_error_handler
    def execute(self, document_path: str) -> ReportInputDTO:
        """Run the complete document analysis pipeline and return aggregated results."""
        paragraphs = self._document_text_port.read_paragraphs(path=document_path)
        document_content = self._extract_content(paragraphs=paragraphs, docx_path=document_path)

        citations = self._citation_extraction_port.extract_citations(docx_path=document_path)
        references, section_type = self._reference_extraction_port.extract_references(
            docx_path=document_path
        )

        author_year_citations = [
            (c.text, c.location, paragraphs[c.location])
            for c in citations
            if c.citation_type == CitationType.AUTHOR_YEAR
        ]
        apa_validation = self._validate_apa(citations=author_year_citations)

        grammar = self._check_grammar(paragraphs=paragraphs)
        classification = self._article_classifier.classify(document_content=document_content)
        quality = self._quality_analyzer.analyze(document_content=document_content)

        effective_type = classification.effective_structure_type
        has_references = len(references) > 0
        structure = self._validate_structure(
            document_content=document_content,
            article_type=effective_type,
            has_references=has_references,
        )

        try:
            section_name = SectionName(section_type)
        except ValueError:
            section_name = SectionName.REFERENCES

        matched_citations = self._citation_matcher.match_citations_to_references(
            citations=citations,
            references=references,
            section_type=section_name,
        )

        eumic_violations = self._document_format_inspection_port.inspect(
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

    def _extract_content(self, paragraphs: list[str], docx_path: str) -> DocumentContentDTO:
        base = self._content_extraction_port.extract(paragraphs=paragraphs, docx_path=docx_path)
        try:
            counts = self._character_count_port.count(docx_path=docx_path)
        except CharacterCountUnavailable:
            return base
        if counts is None:
            return base
        return replace(
            base,
            word_count=counts.word_count,
            char_count=counts.char_count,
            paragraph_count=counts.paragraph_count,
        )

    def _validate_apa(self, citations: list[tuple[str, int, str]]) -> ApaValidationResultDTO:
        if not citations:
            return ApaValidationResultDTO(is_valid=True, violation_count=0, violations=[])
        violations = self._apa_validator.validate_all_citations(citations=citations)
        count = len(violations)
        return ApaValidationResultDTO(
            is_valid=(count == 0), violation_count=count, violations=violations
        )

    def _check_grammar(self, paragraphs: list[str]) -> GrammarCheckResultDTO:
        errors = self._grammar_check_port.check(paragraphs=paragraphs)
        level = GrammarScoreLevel.from_error_count(error_count=len(errors))
        return GrammarCheckResultDTO(score=level.score, feedback=level.feedback, errors=errors)

    def _validate_structure(
        self,
        document_content: DocumentContentDTO,
        article_type: ArticleType,
        has_references: bool,
    ) -> StructureValidationResultDTO:
        if not document_content.paragraphs:
            raise DocumentEmpty

        _, missing = self._structure_validator.validate(
            document_content=document_content, article_type=article_type
        )
        missing = [s for s in missing if s != SectionName.DEVELOPMENT]
        if has_references:
            missing = [s for s in missing if s != SectionName.REFERENCES]

        return StructureValidationResultDTO(
            is_valid=len(missing) == 0,
            missing_sections=list(missing),
        )
