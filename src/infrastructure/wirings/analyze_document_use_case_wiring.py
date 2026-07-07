from dotenv import load_dotenv

from src.application.analyze_document_use_case import AnalyzeDocumentUseCase
from src.domain.citation.apa_validator import ApaValidator
from src.domain.citation.citation_matcher import CitationMatcher
from src.domain.classification.article_classification_response_parser import (
    ArticleClassificationResponseParser,
)
from src.domain.classification.article_classification_text_sampler import (
    ArticleClassificationTextSampler,
)
from src.domain.classification.article_classifier import ArticleClassifier
from src.domain.classification.article_size_classifier import ArticleSizeClassifier
from src.domain.classification.classification_rule_table import ClassificationRuleTable
from src.domain.classification.imryd_signal_detector import ImrydSignalDetector
from src.domain.classification.methodological_vocabulary_detector import (
    MethodologicalVocabularyDetector,
)
from src.domain.classification.reference_signal_detector import ReferenceSignalDetector
from src.domain.document.character_count_port import CharacterCountPort
from src.domain.document.citation_extraction_port import CitationExtractionPort
from src.domain.document.content_extraction_port import ContentExtractionPort
from src.domain.document.document_format_inspection_port import DocumentFormatInspectionPort
from src.domain.document.document_text_port import DocumentTextPort
from src.domain.document.reference_extraction_port import ReferenceExtractionPort
from src.domain.dtos.article_size_thresholds_dto import ArticleSizeThresholdsDTO
from src.domain.dtos.quality_level_thresholds_dto import QualityLevelThresholdsDTO
from src.domain.grammar.grammar_check_port import GrammarCheckPort
from src.domain.ports.llm_generator_port import LlmGeneratorPort
from src.domain.quality.quality_analyzer import QualityAnalyzer
from src.domain.quality.quality_level_resolver import QualityLevelResolver
from src.domain.quality.quality_response_parser import QualityResponseParser
from src.domain.quality.quality_text_sampler import QualityTextSampler
from src.domain.recommendation.recommendation_builder import RecommendationBuilder
from src.domain.structure.structure_validator import StructureValidator
from src.infrastructure.adapters.document.docx_citation_adapter import DocxCitationAdapter
from src.infrastructure.adapters.document.docx_eumic_adapter import DocxEumicAdapter
from src.infrastructure.adapters.document.docx_reference_adapter import DocxReferenceAdapter
from src.infrastructure.adapters.document.docx_text_adapter import DocxTextAdapter
from src.infrastructure.adapters.document.paragraph_content_adapter import ParagraphContentAdapter
from src.infrastructure.adapters.document.win32com_word_count_adapter import (
    Win32ComWordCountAdapter,
)
from src.infrastructure.adapters.grammar.language_tool_adapter import LanguageToolAdapter
from src.infrastructure.adapters.llm_generator.ollama_generator_adapter import (
    OllamaGeneratorAdapter,
)
from src.infrastructure.env_config import EnvConfig
from src.infrastructure.resources.prompts.classification import (
    PROMPTS_DIR as CLASSIFICATION_PROMPTS_DIR,
)
from src.infrastructure.resources.prompts.quality import PROMPTS_DIR as QUALITY_PROMPTS_DIR
from src.infrastructure.resources.text_resource_loader import read_text_resource

load_dotenv()


class AnalyzeDocumentUseCaseWiring:
    """Composition root for the full document analysis pipeline."""

    def __init__(self) -> None:
        self._llm_generator_instance: LlmGeneratorPort | None = None
        self._env_config_instance: EnvConfig | None = None

    def create_use_case(self) -> AnalyzeDocumentUseCase:
        return AnalyzeDocumentUseCase(
            document_text_port=self._get_document_text_port(),
            content_extraction_port=self._get_content_extraction_port(),
            character_count_port=self._get_character_count_port(),
            citation_extraction_port=self._get_citation_extraction_port(),
            reference_extraction_port=self._get_reference_extraction_port(),
            grammar_check_port=self._get_grammar_check_port(),
            document_format_inspection_port=self._get_document_format_inspection_port(),
            apa_validator=self._get_apa_validator(),
            article_classifier=self._get_article_classifier(),
            quality_analyzer=self._get_quality_analyzer(),
            structure_validator=self._get_structure_validator(),
            citation_matcher=self._get_citation_matcher(),
            recommendation_builder=self._get_recommendation_builder(),
        )

    def _get_env_config(self) -> EnvConfig:
        if self._env_config_instance is None:
            self._env_config_instance = EnvConfig()
        return self._env_config_instance

    def _get_document_text_port(self) -> DocumentTextPort:
        return DocxTextAdapter()

    def _get_content_extraction_port(self) -> ContentExtractionPort:
        return ParagraphContentAdapter()

    def _get_character_count_port(self) -> CharacterCountPort:
        return Win32ComWordCountAdapter()

    def _get_citation_extraction_port(self) -> CitationExtractionPort:
        env_config = self._get_env_config()
        return DocxCitationAdapter(
            document_text_port=self._get_document_text_port(),
            max_author_name_length=env_config.citation_max_author_name_length,
        )

    def _get_reference_extraction_port(self) -> ReferenceExtractionPort:
        return DocxReferenceAdapter(document_text_port=self._get_document_text_port())

    def _get_grammar_check_port(self) -> GrammarCheckPort:
        env_config = self._get_env_config()
        return LanguageToolAdapter(max_replacements=env_config.grammar_max_replacements)

    def _get_document_format_inspection_port(self) -> DocumentFormatInspectionPort:
        return DocxEumicAdapter()

    def _get_apa_validator(self) -> ApaValidator:
        return ApaValidator()

    def _get_citation_matcher(self) -> CitationMatcher:
        return CitationMatcher()

    def _get_structure_validator(self) -> StructureValidator:
        env_config = self._get_env_config()
        return StructureValidator(max_header_length=env_config.structure_max_header_length)

    def _get_recommendation_builder(self) -> RecommendationBuilder:
        return RecommendationBuilder(settings=self._get_env_config().get_recommendation_settings())

    def _get_article_classifier(self) -> ArticleClassifier:
        env_config = self._get_env_config()
        return ArticleClassifier(
            llm_generator=self._get_llm_generator(),
            signal_detector=ImrydSignalDetector(),
            article_size_classifier=self._get_article_size_classifier(),
            text_sampler=ArticleClassificationTextSampler(),
            response_parser=ArticleClassificationResponseParser(),
            signal_prompt_template=read_text_resource(
                directory=CLASSIFICATION_PROMPTS_DIR, filename="s4_s5_s6_signal_prompt.txt"
            ),
            temperature=env_config.article_classifier_temperature,
            num_predict=env_config.article_classifier_num_predict,
            methodological_vocabulary_detector=MethodologicalVocabularyDetector(),
            reference_signal_detector=ReferenceSignalDetector(),
            rule_table=ClassificationRuleTable(),
        )

    def _get_article_size_classifier(self) -> ArticleSizeClassifier:
        return ArticleSizeClassifier(thresholds=self._get_article_size_thresholds())

    def _get_article_size_thresholds(self) -> ArticleSizeThresholdsDTO:
        env_config = self._get_env_config()
        return ArticleSizeThresholdsDTO(
            short_min_chars=env_config.article_size_short_min_chars,
            short_max_chars=env_config.article_size_short_max_chars,
            undefined_min_chars=env_config.article_size_undefined_min_chars,
            undefined_max_chars=env_config.article_size_undefined_max_chars,
            long_min_chars=env_config.article_size_long_min_chars,
            long_max_chars=env_config.article_size_long_max_chars,
        )

    def _get_quality_analyzer(self) -> QualityAnalyzer:
        return QualityAnalyzer(
            llm_generator=self._get_llm_generator(),
            text_sampler=self._get_quality_text_sampler(),
            response_parser=QualityResponseParser(),
            clarity_coherence_prompt_template=read_text_resource(
                directory=QUALITY_PROMPTS_DIR, filename="clarity_coherence_prompt.txt"
            ),
            argumentation_conclusions_prompt_template=read_text_resource(
                directory=QUALITY_PROMPTS_DIR, filename="argumentation_conclusions_prompt.txt"
            ),
            resolver=self._get_quality_level_resolver(),
        )

    def _get_quality_level_resolver(self) -> QualityLevelResolver:
        return QualityLevelResolver(thresholds=self._get_quality_level_thresholds())

    def _get_quality_level_thresholds(self) -> QualityLevelThresholdsDTO:
        env_config = self._get_env_config()
        return QualityLevelThresholdsDTO(
            excellent_threshold=env_config.quality_level_excellent_threshold,
            good_threshold=env_config.quality_level_good_threshold,
            acceptable_threshold=env_config.quality_level_acceptable_threshold,
            needs_improvement_threshold=env_config.quality_level_needs_improvement_threshold,
        )

    def _get_quality_text_sampler(self) -> QualityTextSampler:
        env_config = self._get_env_config()
        return QualityTextSampler(
            min_sample_word_count=env_config.quality_min_sample_word_count,
            text_sample_character_limit=env_config.quality_text_sample_character_limit,
        )

    def _get_llm_generator(self) -> LlmGeneratorPort:
        if self._llm_generator_instance is None:
            env_config = self._get_env_config()
            self._llm_generator_instance = OllamaGeneratorAdapter(
                model_name=env_config.ollama_model_name, base_url=env_config.ollama_base_url
            )
        return self._llm_generator_instance
